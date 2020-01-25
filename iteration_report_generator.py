import requests, json, sys, smtplib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from tabulate import tabulate

configuration = {
  'smtp_host': '',
  'sender': '',
  'recipients': '',
  'project_name': '',
  'iteration_name': '',
  'typeoid': '',
  'api_token': '',
  'api_url_base': 'https://rally1.rallydev.com/slm/webservice/v2.x'
}

with open("settings.cfg") as f:
    l = [line.split("=") for line in f.readlines()]
    configuration = {key.strip(): value.strip() for key, value in l}

headers = {'Content-Type': 'application/json',
        'Authorization': 'Bearer {0}'.format(configuration['api_token'])}


def send_email(text,html):
    message = MIMEMultipart(
        "alternative", None, [MIMEText(text), MIMEText(html,'html')])

    message['Subject'] = "Rally Bot - {} Iteration Status for {} ".format(configuration['project_name'],configuration['iteration_name'])
    message['From'] = configuration['sender']
    message['To'] = configuration['recipients']
    smtp_server = smtplib.SMTP(configuration['smtp_host'])
    smtp_server.ehlo()
    smtp_server.starttls()
    smtp_server.sendmail(configuration['sender'], configuration['recipients'], message.as_string())
    smtp_server.quit()
    print("Mail Sent")

def get_project_details(): 


    api_url = '{0}/projects'.format(configuration['api_url_base'])

    params = {
        'query' : '(Name = "{}")'.format(configuration['project_name']),
        'fetch' : 'ObjectID,Workspace,Workspace.ObjectId'
    }
    response = requests.get(api_url, headers=headers, params=params)


    if response.status_code == 200:
        result = json.loads(response.content.decode('utf-8'))
        if result['QueryResult']['TotalResultCount'] > 0:
            return {
                'projectID': result['QueryResult']['Results'][0]['ObjectID'] , 
                'workspace': result['QueryResult']['Results'][0]['Workspace']['ObjectID'], 
                'projectRefId': result['QueryResult']['Results'][0]['_refObjectUUID']
            }
    else:
        return {'projectID': None, 'workspace': None}

def get_project_tag_id(): 


    api_url = '{0}/tags'.format(configuration['api_url_base'])

    params = {
        'query' : '(Name = "{}")'.format(configuration['tag_name']),
        'fetch' : 'ObjectID'
    }
    response = requests.get(api_url, headers=headers, params=params)


    if response.status_code == 200:
        result = json.loads(response.content.decode('utf-8'))
        if result['QueryResult']['TotalResultCount'] > 0:
            return result['QueryResult']['Results'][0]['ObjectID']
    else:
        return None


def get_iteration(): 

    iteration = {}
    api_url = '{0}/iteration'.format(configuration['api_url_base'])

    params = {
        'query' : '(Name = "{}")'.format(configuration['iteration_name']),
        'fetch' : 'ObjectID,StartDate,EndDate'
    }
    response = requests.get(api_url, headers=headers, params=params)


    if response.status_code == 200:
        result = json.loads(response.content.decode('utf-8'))
        if result['QueryResult']['TotalResultCount'] > 0:
            iteration['Name'] = result['QueryResult']['Results'][0]['_refObjectName']
            iteration['IterationID'] = result['QueryResult']['Results'][0]['ObjectID']
            iteration['StartDate'] = datetime.strptime(result['QueryResult']['Results'][0]['StartDate'], "%Y-%m-%dT%H:%M:%S.%fZ").strftime('%d-%b-%Y')
            iteration['EndDate'] = datetime.strptime(result['QueryResult']['Results'][0]['EndDate'], "%Y-%m-%dT%H:%M:%S.%fZ").strftime('%d-%b-%Y')
            return iteration
    else:
        return None


def get_iterationstatus(project_config): 

    api_url = '{0}/artifact'.format(configuration['api_url_base'])
    

    query = '(((TypeDefOid = {0}) AND (Tags CONTAINS "/tag/{1}")) AND (Iteration.Name = "{2}"))'.format(configuration['typeoid'],get_project_tag_id(),configuration['iteration_name'])  if configuration['tag_name'] != '' else '(Iteration.Name = "{}")'.format(configuration['iteration_name'])
    params = {
        'compact': 'false',
        'includePermissions': 'true',
        'projectScopeUp': 'false',
        'projectScopeDown': 'true',
        'showHiddenFieldsForVersionedAlias': 'true',
        'start': 1,
        'pagesize': 500,
        'order': 'DragAndDropRank DESC',
        'types' : 'HierarchicalRequirement,Defect,TestSet,DefectSuite',
        'project': '/project/{}'.format(project_config['projectRefId']),
        'query' : query,
        'fetch' : 'TestFolder,DisplayColor,TestCase,Requirement,DefectStatus,Ready,Actuals,ToDo,DirectChildrenCount,Release,Estimate,Parent,Name,Blocked,ScheduleStatePrefix,TaskIndex,Tasks,TestSet,BlockedReason,State,TestCases,TaskActualTotal,PlanEstimate,Owner,TaskRemainingTotal,TaskEstimateTotal,FormattedID,Defects:summary[State],Project,ScheduleState,PortfolioItem,Iteration,WorkProduct,DragAndDropRank,Children,sum:[TaskActualTotal,TaskRemainingTotal,PlanEstimate,TaskEstimateTotal]'
    }
    response = requests.get(api_url, headers=headers, params=params)


    if response.status_code == 200:
        result = json.loads(response.content.decode('utf-8'))
        if result['QueryResult']['TotalResultCount'] > 0:
            return {'iterationResult': result['QueryResult']['Results'], 'iterationSummary': result['QueryResult']['Sums'] }
    else:
        return {'iterationResult': None}

def generate_report(report):

    iteration = get_iteration()
    
    text = """
    Iteration Status for : {project_name}
    Iteration : {Name}
    Start date : {StartDate}
    End date : {EndDate}
    {summary_table}

    {stories_table}

    Regards,

    Rally Bot"""

    html = """
    <html>
    <style> 
    table {{
    font-family: "Helvetica, sans-serif;
    border: 1px solid black;
    border-collapse: collapse;
    width: 100%;
    }}

    td, th {{
    border: 1px solid #black;
    padding: 8px;
    }}

    tr:nth-child(even){{background-color: #f2f2f2;}}

    th {{
    padding-top: 12px;
    padding-bottom: 12px;
    text-align: left;
    background-color: #0398fc;
    color: white;
    }}
    </style>
    <body>
    <h1><u>Iteration Status for : {project_name}</u></h1>
    <h3>Iteration : {Name} </h3>
    <h3>Start date  : {StartDate} </h3>
    <h3>End date  : {EndDate} </h3>
    <br/><br/>
    {summary_table}
    <br/><br/>
    {stories_table}
    <br/><br/>
    <p>Regards,</p>
    <p>Rally Bot</p>
    </body></html>
    """
    data_header = ["ID", 'Project Name', "Name", "Owner", "State", "Blocked", "Points", "Total hrs", "Remaining hrs"]
    data = []
    data.append(data_header)
    summary = {'In-Progress': 0, 'Accepted': 0, 'Completed': 0, 'Defined': 0, 'Total': 0}
    for result in report['iterationResult']:
        summary[result['ScheduleState']] =  summary[result['ScheduleState']] + result['PlanEstimate']
        summary['Total'] = summary['Total'] + result['PlanEstimate']
        data.append([result['FormattedID'],result['Project']['_refObjectName'], result['Name'], result['Owner']['_refObjectName'], result['ScheduleState'], result['Blocked'], result['PlanEstimate'], result['TaskEstimateTotal'], result['TaskActualTotal']])
    text = text.format(stories_table=tabulate(data, headers="firstrow", tablefmt="grid"),
            summary_table=tabulate(summary.items(),headers=["State", "Story Points"],tablefmt="grid"),
            project_name=configuration['project_name'],
            **iteration)
    html = html.format(stories_table=tabulate(data, headers="firstrow", tablefmt="html"),
            summary_table=tabulate(summary.items(),headers=["State", "Story Points"],tablefmt="html"),
            project_name=configuration['project_name'],
            **iteration)
        
    
    send_email(text,html)

def main():
    project_config = get_project_details()
    
    if project_config == None:
        print("Project Not found, exiting")
        sys.exit(0)
    report = get_iterationstatus(project_config)
    generate_report(report)




if __name__ == '__main__':
    main()
    sys.exit(0)
