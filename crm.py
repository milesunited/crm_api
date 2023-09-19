import requests
import json
import openpyxl

class GraphQLClient:
    def __init__(self, url, token):
        self.url = url
        self.headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {token}"
        }

    def execute_query(self, query):
        data = json.dumps({"query": query})
        response = requests.post(self.url, headers=self.headers, data=data)
        return response.json()

def main():
    api_url = "https://api.fireflies.ai/graphql/"
    bearer_token = "xxxxxxxxxxxxxxxxxxxxxxx"
    query = "{ transcripts { title date } }"

    client = GraphQLClient(api_url, bearer_token)
    result = client.execute_query(query)

    # Create or load an Excel workbook
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # Write headers
    sheet['A1'] = "Title"
    sheet['B1'] = "Date"

    # Write data to the sheet
    for index, entry in enumerate(result):
        sheet[f'A{index+1}'] = entry[1]
        sheet[f'B{index+1}'] = entry[2]

    # Save the workbook to a file
    workbook.save('transcripts.xlsx')

if __name__ == "__main__":
    main()
