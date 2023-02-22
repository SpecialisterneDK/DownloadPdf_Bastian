import pandas as pd
import requests
import os
import os.path
import glob
import concurrent.futures
import xlsxwriter

# check if the downloaded.xlsx file already exists
if os.path.exists("downloaded.xlsx"):
    # read in the existing downloaded.xlsx file and update the "Downloaded" column
    existing_df = pd.read_excel("downloaded.xlsx", sheet_name='Downloaded', index_col='BRnum')
    df2 = pd.read_excel("GRI_2017_2020.xlsx", sheet_name=0, index_col='BRnum')
    df2['Downloaded'] = existing_df['Downloaded']
    df2['Pdf_URL_AL'] = existing_df['Pdf_URL_AL']
    df2['Pdf_URL_AM'] = existing_df['Pdf_URL_AM']
else:
    # create a new df2 DataFrame with empty columns
    df2 = pd.read_excel("GRI_2017_2020.xlsx", sheet_name=0, index_col='BRnum')
    df2['Downloaded'] = ""
    df2['Pdf_URL_AL'] = ""
    df2['Pdf_URL_AM'] = ""
    df2 = df2[df2.index.notnull()]

# create a new downloaded.xlsx file
workbook = xlsxwriter.Workbook("downloaded.xlsx")
worksheet = workbook.add_worksheet('Downloaded')

worksheet.write('A1', 'BRnum')
worksheet.write('B1', 'Pdf_URL_AL')
worksheet.write('C1', 'Pdf_URL_AM')
worksheet.write('D1', 'Downloaded')

# specify output folder (in this case it moves one folder up and saves in the script output folder)
pth = 'TextMining'

# specify path for existing downloads
dwn_pth = 'TextMining/dwn/'

# filter out rows with no URL
non_empty = df2.Pdf_URL.notnull()
df2 = df2[non_empty]

# create download directory if it doesn't exist
os.makedirs(dwn_pth, exist_ok=True)


# define the download function
def download_file(j):
    savefile = os.path.join(dwn_pth, f"{j}.pdf")
    if os.path.exists(savefile):
        print(f"{savefile} already exists, skipping")
        return f"{savefile} already exists, skipping"
    try:
        url = df2.at[j, 'Pdf_URL']
        response = requests.get(url, stream=True, timeout=30)
        response.raise_for_status()
        # check if the response header indicates a PDF file
        if response.headers.get('content-type') == 'application/pdf':
            with open(savefile, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            print(f"Downloaded {savefile}")
            # update the "Downloaded" and "Pdf_URL_AL" columns in the df2 DataFrame
            df2.at[j, "Downloaded"] = "Yes"
            df2.at[j, "Pdf_URL_AL"] = url
            return f"Downloaded {savefile} using Pdf_URL"
        else:
            # update the "Downloaded" column in the df2 DataFrame
            df2.at[j, "Downloaded"] = "No"
            raise ValueError("Invalid URL: Pdf_URL")
    except:
        try:
            url = df2.at[j, 'Report Html Address']
            response = requests.get(url, stream=True, timeout=30)
            response.raise_for_status()
            # check if the response header indicates a PDF file
            if response.headers.get('content-type') == 'application/pdf':
                with open(savefile, 'wb') as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        f.write(chunk)
                print(f"Downloaded {savefile} from Report Html Address column")
                # update the "Downloaded" and "Pdf_URL_AM" columns in the df2 DataFrame
                df2.at[j, "Downloaded"] = "Yes"
                df2.at[j, "Pdf_URL_AM"] = url
                return f"Downloaded {savefile} using Report Html Address column"
            else:
                # update the "Downloaded" column in the df2 DataFrame
                df2.at[j, "Downloaded"] = "No"
                raise ValueError("Invalid URL: Report Html Address")
        except:
            # update the "Downloaded" column in the df2 DataFrame
            df2.at[j, "Downloaded"] = "No"
            return None



# download the files concurrently
with concurrent.futures.ThreadPoolExecutor() as executor:
    results = executor.map(download_file, df2.index)

# print out the successful downloads
for result in results:
    if result is not None:
        print(result)

# write the updated df2 DataFrame to the "Downloaded" worksheet using pandas ExcelWriter
writer = pd.ExcelWriter('downloaded.xlsx', engine='xlsxwriter')
df2[['Pdf_URL_AL', 'Pdf_URL_AM', 'Downloaded']].to_excel(writer, sheet_name='Downloaded')
writer.save()

print("files have been downloaded!")
