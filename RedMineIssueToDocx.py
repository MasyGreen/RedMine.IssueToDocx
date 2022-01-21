import os
import requests  # request img from web
import shutil  # save img locally
import uuid

import keyboard
import lxml
from bs4 import BeautifulSoup
from colorama import Fore, Style, init, AnsiToWin32
from htmldocx import HtmlToDocx
import re
from datetime import datetime

from redminelib import Redmine
import configparser


def ReadConfig(configfilepath):
    print(f'{Fore.YELLOW}Start ReadConfig')

    if os.path.exists(configfilepath):
        config = configparser.ConfigParser()
        config.read(configfilepath)
        config.sections()

        global glhost
        glhost = str(config.get("Settings", "host"))

        global glapikey
        glapikey = config.get("Settings", "apikey")

        global glissuesid
        glissuesid = config.get("Settings", "issuesid")

        return True
    else:
        print(f'{Fore.YELLOW}Start create_config')
        config = configparser.ConfigParser()
        config.add_section("Settings")

        config.set("Settings", "host", r'http://192.168.1.1')
        config.set("Settings", "apikey", r'dq3inqgnqe8igqngninkkvekmviewrgir9384')
        config.set("Settings", "issuesid", r'1677;318')

        with open(configfilepath, "w") as config_file:
            config.write(config_file)

        print(f'{Fore.GREEN}Create config: {Fore.BLUE}{configfilepath},{config_file}')
        return False


def writedocx(issueDescription, filename):
    print(f'{Fore.YELLOW}Write *.docx: {filename}')

    new_parser = HtmlToDocx()
    docx = new_parser.parse_html_string(issueDescription)
    docx.save(filename)


def main():
    redmine = Redmine(glhost, key=glapikey)

    print(f'{Fore.CYAN}{currentDirectory=}')

    now = datetime.now()
    downloadDirectory = os.path.join(currentDirectory, str(f'{now.strftime("%Y-%m-%d")}-{now.strftime("%H-%M-%S")}'))
    if not os.path.exists(downloadDirectory):
        os.makedirs(downloadDirectory)
        print(f"{Fore.CYAN}{downloadDirectory=}")
    print()
    print(f'{Fore.CYAN}===========================================================================')
    indx = 0
    for issueid in issueidlist:
        if issueid.isdigit():
            indx = indx + 1
            print(f'{Fore.GREEN}Process {indx}: {issueid=}')
            # 1 Get RedMine Issue Description
            issue = redmine.issue.get(issueid, include=[])
            issueDescription = issue.description

            # Process image
            # 2.1 Collect Scr list
            soup = BeautifulSoup(issueDescription, "lxml")
            imgScrs = soup.findAll("img")

            # 2.2 Download image from Src
            for item in imgScrs:
                imgScr = item.get("src")
                pathImg = f'{glhost}/{imgScr}'
                print(f'{Fore.BLUE}{pathImg=}')

                xfilename, xfile_extension = os.path.splitext(pathImg)
                newImgName = f'{uuid.uuid4()}{xfile_extension}'
                newImPath = os.path.join(downloadDirectory, newImgName)
                print(f'{Fore.BLUE}Download Img: {pathImg} -> {newImPath=}')
                dowloadLink = f"{pathImg}?key={glapikey}"
                print(f'{dowloadLink=}')
                print(f'{Fore.MAGENTA}Replace: {str(imgScr)} to {str(newImPath)}')

                # Download img
                res = requests.get(dowloadLink, stream=True)
                if res.status_code == 200:
                    with open(newImPath, 'wb') as f:
                        shutil.copyfileobj(res.raw, f)

                # Replace Img link to local path
                issueDescription = issueDescription.replace(str(imgScr), str(newImPath))
                print()

            # 3 Create *.docx
            issuefilename = os.path.join(downloadDirectory, f'Issue - {issueid}.docx')
            writedocx(issueDescription, issuefilename)

            print(f'{Fore.CYAN}_________________________________________________________')



if __name__ == "__main__":
    print(f"{Fore.CYAN}Last update: Cherepanov Maxim masygreen@gmail.com (c), 01.2022")
    print(f"{Fore.CYAN}Convert file *.html to *.docx")

    currentDirectory = os.getcwd()
    configfilepath = os.path.join(currentDirectory, 'config.cfg')

    if ReadConfig(configfilepath):
        issueidlist = glissuesid.split(';')
        print(f'{issueidlist=}')

        main()
    else:
        print(f'{Fore.RED}Pleas edit default Config value: {Fore.BLUE}{configfilepath}')
        keyboard.wait("space")
