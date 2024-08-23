import os
import google.auth
from googleapiclient.discovery import build
from openpyxl import Workbook

# Authenticate and initialize the Google Compute Engine API
def get_compute_service():
    credentials, project = google.auth.default()
    return build('compute', 'v1', credentials=credentials), project

# Fetch all instances in the project
def list_instances(compute, project):
    instances = []
    request = compute.instances().aggregatedList(project=project)
    
    while request is not None:
        response = request.execute()

        for zone, instances_scoped_list in response['items'].items():
            if 'instances' in instances_scoped_list:
                for instance in instances_scoped_list['instances']:
                    instance_data = {
                        "PROJECT_ID": project,
                        "CREATE_TIME": instance.get("creationTimestamp"),
                        "INSTANCE_NAME": instance.get("name"),
                        "MACHINE_TYPE": instance.get("machineType").split('/')[-1],
                        "STATUS": instance.get("status"),
                        "ZONE": instance.get("zone").split('/')[-1],
                        "INSTANCE_ID": instance.get("id"),
                        "SERVICE_ACCOUNT": instance.get("serviceAccounts", [{}])[0].get("email"),
                        "TAGS": ','.join(instance.get("tags", {}).get("items", [])),
                        "SPOT": instance.get("scheduling", {}).get("preemptible", False),
                        "PREEMPTIBLE": instance.get("scheduling", {}).get("preemptible", False)
                    }
                    instances.append(instance_data)
        request = compute.instances().aggregatedList_next(previous_request=request, previous_response=response)

    return instances

# Write instances data to Excel
def write_to_excel(instances, file_name='gcp_instances.xlsx'):
    workbook = Workbook()
    sheet = workbook.active

    headers = ["PROJECT_ID", "CREATE_TIME", "INSTANCE_NAME", "MACHINE_TYPE", "STATUS", "ZONE", "INSTANCE_ID", "SERVICE_ACCOUNT", "TAGS", "SPOT", "PREEMPTIBLE"]
    sheet.append(headers)

    for instance in instances:
        row = [
            instance["PROJECT_ID"],
            instance["CREATE_TIME"],
            instance["INSTANCE_NAME"],
            instance["MACHINE_TYPE"],
            instance["STATUS"],
            instance["ZONE"],
            instance["INSTANCE_ID"],
            instance["SERVICE_ACCOUNT"],
            instance["TAGS"],
            instance["SPOT"],
            instance["PREEMPTIBLE"]
        ]
        sheet.append(row)

    workbook.save(file_name)
    print(f'Data written to {file_name}')

# Main function to fetch and write VM data
def main():
    compute, project = get_compute_service()
    instances = list_instances(compute, project)
    write_to_excel(instances)

if __name__ == "__main__":
    main()
