import requests, json, sys, smtplib
from datetime import datetime, date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from tabulate import tabulate

configuration = {
  'smtp_host': '',
  'sender': '',
  'recipients': '',
  'parent': '',
  'projects': '',
  'typeoid': '',
  'api_token': '',
  'api_url_base': 'https://rally1.rallydev.com/slm/webservice/v2.x'
}

projects = []

with open("configuration.cfg") as f:
    l = [line.split("=") for line in f.readlines()]
    configuration = {key.strip(): value.strip() for key, value in l}

headers = {'Content-Type': 'application/json',
        'Authorization': 'Bearer {0}'.format(configuration['api_token'])}


def send_email(text,html):
    message = MIMEMultipart(
        "alternative", None, [MIMEText(text), MIMEText(html,'html')])

    message['Subject'] = "Rally Bot - {} Iteration Status for {} ".format(configuration['parent'],date.today().strftime("%b-%d-%Y"))
    message['From'] = configuration['sender']
    message['To'] = configuration['recipients']
    smtp_server = smtplib.SMTP(configuration['smtp_host'])
    smtp_server.ehlo()
    smtp_server.starttls()
    smtp_server.sendmail(configuration['sender'], configuration['recipients'], message.as_string())
    smtp_server.quit()
    print("Mail Sent")

def get_projects(): 

    project_configs = configuration['projects'].split(",")
    for project_config in project_configs:
        project_name = project_config.split(':')[0]
        project_iteration = project_config.split(':')[1]
        api_url = '{0}/projects'.format(configuration['api_url_base'])

        params = {
            'query' : '(Name = "{}")'.format(project_name),
            'fetch' : 'ObjectID,Workspace,Workspace.ObjectId'
        }
        response = requests.get(api_url, headers=headers, params=params)


        if response.status_code == 200:
            result = json.loads(response.content.decode('utf-8'))
            if result['QueryResult']['TotalResultCount'] > 0:
                projects.append({
                    'Name': project_name,
                    'projectID': result['QueryResult']['Results'][0]['ObjectID'] , 
                    'workspace': result['QueryResult']['Results'][0]['Workspace']['ObjectID'], 
                    'projectRefId': result['QueryResult']['Results'][0]['_refObjectUUID'],
                    'projectIteration': project_iteration
                })

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


def get_iteration(iterationName): 

    iteration = {}
    api_url = '{0}/iteration'.format(configuration['api_url_base'])

    params = {
        'query' : '(Name = "{}")'.format(iterationName),
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


def get_project_summary(project): 

    api_url = '{0}/hierarchicalrequirement'.format(configuration['api_url_base'])
    

    query = '(((TypeDefOid = {0}) AND (Tags CONTAINS "/tag/{1}")) AND (Iteration.Name = "{2}"))'.format(configuration['typeoid'],get_project_tag_id(),project['projectIteration'])  if configuration['tag_name'] != '' else '(Iteration.Name = "{}")'.format(project['projectIteration'])
    params = {
        'start': 1,
        'pagesize': 500,
        'project': '/project/{}'.format(project['projectRefId']),
        'query' : query,
        'fetch' : 'Tasks:summary[State],sum:[TaskActualTotal,TaskRemainingTotal,PlanEstimate,TaskEstimateTotal]'
    }
    response = requests.get(api_url, headers=headers, params=params)


    if response.status_code == 200:
        result = json.loads(response.content.decode('utf-8'))
        if result['QueryResult']['TotalResultCount'] > 0:
            print('Total Records fetched for {} : {}'.format(project['Name'], result['QueryResult']['TotalResultCount']))
            return {'iterationResult': result['QueryResult']['Results'], 'iterationSummary': result['QueryResult']['Sums'], 'count' : result['QueryResult']['TotalResultCount'] }
    else:
        return {'iterationResult': {}, 'iterationSummary': {}, count: 0}

def generate_report():
    
    text = """
    Iteration Summary :

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
    <h1><u>Iteration Summary</u></h1>
    <br/><br/>
    {summary_table}
    <br/><br/>
    <p>Regards,</p>
    <p>Rally Bot</p>
    </body></html>
    """
    summary_header = ["Application Group", 'Iteration','Assigned User Stories', "Assigned Story Points", "Estimated Hours", "Actual Hours", "% Complete (by Hours)", "#Tasks", "Tasks Completed", "% Complete (by Task)"]
    summary_data = []
    summary_data.append(summary_header)
    task_summary = {'In-Progress': 0, 'Completed': 0, 'Defined': 0, 'Total': 0}
    for project in projects :
        print('Generationg Summary for Project : {} in Iteration {}'.format(project['Name'],project['projectIteration']))
        project_summary =  get_project_summary(project)
        for result in project_summary['iterationResult']:
            for item in task_summary:
                value = result['Summary']['Tasks']['State'][item] if item in result['Summary']['Tasks']['State'] else 0
                task_summary[item] =  task_summary[item] + value
            task_summary['Total'] = task_summary['Total'] +  result['Summary']['Tasks']['Count']
        percent_completed_hrs = '{} %'.format(int((project_summary['iterationSummary']['TaskActualTotal'] / project_summary['iterationSummary']['TaskEstimateTotal']) * 100))
        percent_completed_task = '{} %'.format(int((task_summary['Completed'] / task_summary['Total']) * 100))
        summary_data.append([project['Name'], project['projectIteration'],  project_summary['count'], project_summary['iterationSummary']['PlanEstimate'],project_summary['iterationSummary']['TaskEstimateTotal'], project_summary['iterationSummary']['TaskActualTotal'], percent_completed_hrs, task_summary['Total'], task_summary['Completed'],percent_completed_task])
    text = text.format(summary_table=tabulate(summary_data, headers="firstrow", tablefmt="grid"))
    html = html.format(summary_table=tabulate(summary_data, headers="firstrow", tablefmt="html"))
    #print(summary_data)
    send_email(text,html)

def main():
    get_projects()
    generate_report()


if __name__ == '__main__':
    main()
    sys.exit(0)
   
