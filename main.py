import xmltodict
import os 
import pandas as pd

def getInformation(fileName, valuesList):
    with open('nfs/' + fileName, 'rb') as file: # Open the file
        dicFile = xmltodict.parse(file) # Parse the file to a dictionary
        #print(json.dumps(dicFile, indent=4))
        try:
            if "NFe" in dicFile:
                infoNF = dicFile['NFe']['infNFe'] # Get the information from the file
            else:
                infoNF = dicFile['nfeProc']['NFe']['infNFe'] # Get the information from the file
            numberNF = infoNF['@Id'] # Get the number of the NF
            nameClient = infoNF['emit']['xNome'] # Get the name of the client
            companyName = infoNF['dest']['xNome']# Get the name of the company
            adress = infoNF['dest']['enderDest'] # Get the adress of the company
            if 'vol' in infoNF['transp']:
                weight = infoNF['transp']['vol']['pesoB'] # Get the weight of the product
            else:
                weight = 'Não informado' # If the weight is not informed, set the value to 'Não informado'
            valuesList.append([numberNF,companyName,nameClient, adress, weight]) # Append the values to the list
        except Exception as e:
            print('Error: ' + str(e)) # Print the error
        pass

files = os.listdir('nfs') # List all files in the nfs directory
columns = ['NumberNF', 'Company_Name', 'Name_Client', 'Adress', 'Weight'] # Create the columns for the dataframe
valuesList = [] # Create an empty list to store the values
for file in files: # Loop through all files
    print('File: ' + file) # Print the name of the file
    getInformation(file, valuesList)   # Call the function to get information from the file

table = pd.DataFrame(valuesList, columns=columns) # Create a dataframe with the values 
table.to_excel('NF.xlsx', index=False) # Save the dataframe to a excel file
print(table) # Print the dataframe