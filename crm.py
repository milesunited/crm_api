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
    bearer_token = "d5f23de2-de12-4db6-a806-703be6d78ca1"
    query = "{ transcripts { title date id participants transcript_url duration } }"

    client = GraphQLClient(api_url, bearer_token)
    result = client.execute_query(query)
    #print(result)
#    print(result['data'])

    # Create or load an Excel workbook
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # Write headers
    sheet['A1'] = "Title"
    sheet['B1'] = "Date"
    sheet['C1'] = "ID"
    sheet['D1'] = "People"
    sheet['E1'] = "Transcripts"



    data_ = result['data']['transcripts']
    print(data_[0])
    for val, dat in enumerate(data_):
        print(f"val: {str(val)} :: data: {str(dat)} " )
        sheet[f'A{val+2}'] =str(dat["title"])
        sheet[f'B{val+2}']  =str(dat["date"])        
        sheet[f'C{val+2}']  =str(dat["id"])
        sheet[f'D{val+2}']  =str(dat["participants"])
        sheet[f'E{val+2}']  =str(dat["transcript_url"])   
        sheet[f'F{val+2}']  =str(dat["duration"])



    # Save the workbook to a file
    workbook.save('transcripts.xlsx')

if __name__ == "__main__":
    main()
