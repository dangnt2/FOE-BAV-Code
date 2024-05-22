import PyPDF2
import io
import glob
import re
import os
import pandas as pd
import numpy as np
pd.set_option('display.max_columns', None)
#import playwright sync
from playwright.sync_api import sync_playwright
import time
import pandas as pd
import os
# -------- Mọi thứ đều có công sức từ mọi người ở FOE-BAV, hoàn thiện và phát triển nó -----
# Contact trao đổi : dangnguyen110900@gmail.com
folder_to_extract_names_from_pdf = r".\14-16 Jul 23"
ofp_names = r".\\"

txt_output_folder = r".\14-16 Jul 23"
txt_output_folder_name = folder_to_extract_names_from_pdf.split("\\")[-1]
txt_output_folder_name = "output " + txt_output_folder_name
original_wd = os.getcwd()

# ofp_names = r".\14-16 Jul 23\ofp_names.csv"

def extract_text_from_folder(pdf_folder_path, extracted_text_name, level=3):
    # recursively get all pdf files in the folder with "OFP" in it name
    pdf_folder_path = pdf_folder_path + r"\**\*OFP*.pdf"
    all_files = glob.glob(pdf_folder_path, recursive=True)
    # loop through files to extract data
    df = pd.DataFrame(columns=["ofp_number", "name", "url"])
    for file_num, filez in enumerate(all_files):
        with open(filez, 'rb') as f:
            try:
                # read pdf file
                file_reader = PyPDF2.PdfReader(f)
                # get first page of pdf and extract text
                all_lines = io.StringIO(file_reader.pages[0].extract_text())
            except:
                print("Error in reading file")
                continue
            line_list = []
            for line in all_lines:
                line_list.append(line)
                # print(line)
            # get c_time at line 4. Vd: 0745z
            # đôi khi pdf cách thêm 1 dòn nên phải tìm vào dòng sau của dòng ban đầu.
            row = -2
            for i in range(0, 10):
                try:
                    ofp_number = re.search("(O)(\s)*(F)(\s)*(P)(\s)*((\d)*(\s)*){4}", line_list[row+2]).group()
                    break
                except:
                    row += 1
            try:
                path = f.name.split("\\")
                # get file name. Vd: OFP 8477 BAV88 VVTS TO YMML.pdf
                name = path[-1]
                #change name to OFP 8477 BAV88 VVTS TO YMML
                name = name[:-4]
                #add .txt to name
                name = name + ".txt"

                ofp_number = re.search("(O)(\s)*(F)(\s)*(P)(\s)*((\d)*(\s)*){4}", line_list[row+2]).group().split()
                ofp_number = ''.join(ofp_number)
                ofp_number = ofp_number[-4:]

                #ofp text url
                url = f"https://gold.jetplan.com/jeppesen/jpdcServlet?query=755&planNum={ofp_number}"

                # extract date from string. Vd:".\To Extract\Week 207 09-15 Oct 22\2022.10.15 OFP\BAV88\OFP 8477 BAV88 VVTS TO YMML.pdf"
                # get c_date at line 4. Vd: 03/12/20
            except:
                print("Error in extracting data")
                path_splitted = extracted_text_name.split("\\")
                error_text_path = ""
                for i, e in enumerate(path_splitted):
                    if i < len(path_splitted)-1:
                        error_text_path += f"{e}\\"

                with open(f"{error_text_path}error.txt", "a") as b:
                    b.write(f"{f.name}\n")
                continue
            file_num = str(file_num) + "."
            print(file_num, ofp_number,name,url)
            #add data to df using concat
            df = pd.concat([df, pd.DataFrame({"ofp_number": [ofp_number], "name": [name], "url": [url]})], ignore_index=True)
    return df
            # save dataframe to csv file
            # with open(extracted_text_name, "a") as e:
            #     if os.stat(extracted_text_name).st_size == 0:
            #         e.write("ofp_number,name,url\n")
            #         e.write(f"{ofp_number},{name},{url}\n")
            #     else:
            #         e.write(
            #         f"{ofp_number},{name},{url}\n")

df = extract_text_from_folder(folder_to_extract_names_from_pdf, ofp_names, level=3)

payload = {
    "loginid":"bamboo.read", 
    "password":"OCCbamboo@No01"}

login_url = "https://gold.jetplan.com/jeppesen/jsp/login/login_generic.jsp"


#if df["ofp_number"] is not a string, convert to string
df["ofp_number"] = df["ofp_number"].astype(str)

# add "0" to df["ofp_number"] if length is < 4
df["ofp_number"] = df["ofp_number"].apply(lambda x: x.zfill(4))

# make txt_output_folder as current working directory

os.chdir(txt_output_folder)
if not os.path.exists(txt_output_folder_name):
    os.makedirs(txt_output_folder_name)

txt_output_folder = txt_output_folder + "\\" + txt_output_folder_name + "\\"
# reset working directory
os.chdir(original_wd)
print(txt_output_folder)

print(df)
with sync_playwright() as p:
    browser = p.chromium.launch(headless=False,slow_mo=50)
    context = browser.new_context()
    # Open new page
    page = context.new_page()
    # Go to https://gold.jetplan.com/jeppesen/jsp/login/login_generic.jsp
    page.goto(login_url)
    # fill input[name="loginid"]
    page.fill("input[name=\"loginid\"]", payload["loginid"])
    # fill input[name="password"]
    page.fill("input[name=\"password\"]", payload["password"])
    # click button type="submit"
    page.click("button[type=\"submit\"]")
    time.sleep(0)
    # for row in df, request_url = df["url"], txt = f"https://gold.jetplan.com/jeppesen/DocServlet?fn0=%2Ftemp%2FsaveAs&fn1=BAMBOO2_{df["ofp_number"]}&fn2=txt", save page body to file with name = df["name"]
    for index, row in df.iterrows():
        # if index less than 443, continue
        # if index < 443:
            # continue
        request_url = row["url"]
        txt = f"https://gold.jetplan.com/jeppesen/DocServlet?fn0=%2Ftemp%2FsaveAs&fn1=BAMBOO2_{row['ofp_number']}&fn2=txt"
        page.goto(request_url)
        # time.sleep(1)
        page.goto(txt)
        ####
        ####
        ###
        name = txt_output_folder + row["name"]
        ####
        ####
        ####
        with open(name, "w") as f:
            f.write(page.inner_text("body"))
        # print to console what file is being saved with current time and index

        print(f" {index}. Saved {row['name']}  -   {time.strftime('%H:%M:%S')}")
    
    
    
    # ---------------------