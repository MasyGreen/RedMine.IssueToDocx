import configparser
import os
import shutil  # save img locally
import uuid
from datetime import datetime

import keyboard
import requests  # request img from web
from bs4 import BeautifulSoup
from colorama import Fore
from htmldocx import HtmlToDocx
from redminelib import Redmine


def ReadConfig(filepath):
    print(f'{Fore.YELLOW}Start ReadConfig')

    if os.path.exists(filepath):
        config = configparser.ConfigParser()
        config.read(filepath, "utf8")
        config.sections()

        global glhost
        glhost = config.has_option("Settings", "host") and config.get("Settings", "host") or None

        global glapikey
        glapikey = config.has_option("Settings", "apikey") and config.get("Settings", "apikey") or None

        global glissuesid
        glissuesid = config.has_option("Settings", "issuesid") and config.get("Settings", "issuesid") or None

        global glcombine
        varstr = config.has_option("Settings", "combine") and config.get("Settings", "combine") or None
        if varstr == 'true':
            glcombine = True
        else:
            glcombine = False

        global glsaveimg
        varstr = config.has_option("Settings", "saveimg") and config.get("Settings", "saveimg") or None
        if varstr == 'true':
            glsaveimg = True
        else:
            glsaveimg = False

        global gliswiki
        varstr = config.has_option("Settings", "iswiki") and config.get("Settings", "iswiki") or None
        if varstr == 'true':
            gliswiki = True
        else:
            gliswiki = False

        return True
    else:
        print(f'{Fore.YELLOW}Start create_config')
        config = configparser.ConfigParser()
        config.add_section("Settings")

        config.set("Settings", "host", 'http://192.168.1.1')
        config.set("Settings", "apikey", 'dq3inqgnqe8igqngninkkvekmviewrgir9384')
        config.set("Settings", "issuesid", '1677;318')
        config.set("Settings", "saveimg", 'false')
        config.set("Settings", "combine", 'false')
        config.set("Settings", "iswiki", 'false')

        with open(filepath, "w") as config_file:
            config.write(config_file)

        print(f'{Fore.GREEN}Create config: {Fore.BLUE}{filepath},{config_file}')
        return False


def WriteDocx(issueDescription, filename):
    print(f'{Fore.YELLOW}Write *.docx: {filename}')

    new_parser = HtmlToDocx()
    docx = new_parser.parse_html_string(issueDescription)
    docx.save(filename)


def DownloadIMG(description, imgScrs, downloadDirectory, imagesavelist):
    for item in imgScrs:
        imgScr = item.get("src")
        pathImg = f'{glhost}/{imgScr}'
        print(f'{Fore.BLUE}{pathImg=}')

        xfilename, xfile_extension = os.path.splitext(pathImg)
        newImgName = f'{uuid.uuid4()}{xfile_extension}'
        newImPath = os.path.join(downloadDirectory, newImgName)
        imagesavelist.append(newImPath)  # save list download Img file
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
        description = description.replace(str(imgScr), str(newImPath))

        return description


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

    imagesavelist = []
    issueCombineDescription = ''
    wikiCombineDescription = ''
    for issueid in issueidlist:
        print(f'{issueid=}')
        if issueid != "":
            indx = indx + 1
            print(f'{Fore.GREEN}Process {indx}: {issueid=}')
            print(f'{gliswiki=}')

            if not gliswiki:
                if issueid.isdigit():
                    print(f'Get RedMine Issue Description: {issueid=}')

                    # 1 Get RedMine Issue Description
                    issue = redmine.issue.get(issueid, include=[])
                    issueDescription = issue.description

                    # Process image
                    # 2.1 Collect Scr list
                    soup = BeautifulSoup(issueDescription, "lxml")
                    imgScrs = soup.findAll("img")

                    # 2.2 Download image from Src
                    issueDescription = DownloadIMG(issueDescription, imgScrs, downloadDirectory, imagesavelist)
                    print()

                    # 3 Create *.docx
                    if not glcombine:
                        exportfilename = os.path.join(downloadDirectory, f'Issue - {issueid}.docx')
                        WriteDocx(issueDescription, exportfilename)
                    else:
                        issueDescription = f'<h1>Сохранение: Issue {issueid}</h1>{issueDescription}'
                        issueCombineDescription = issueCombineDescription + issueDescription
            else:
                print(f'Get RedMine Wiki Description: {issueid=}')

                issueiditem = issueid.split('/')
                if len(issueiditem) == 3:
                    projectid = issueiditem[0]
                    wikiname = issueiditem[2]

                    # 1 Get RedMine Wiki Description
                    wiki = redmine.wiki_page.get(wikiname, project_id=projectid, include=[])
                    wikiDescription = wiki.text

                    # Process image
                    # 2.1 Collect Scr list
                    soup = BeautifulSoup(wikiDescription, "lxml")
                    imgScrs = soup.findAll("img")

                    # 2.2 Download image from Src
                    wikiDescription = DownloadIMG(wikiDescription, imgScrs, downloadDirectory, imagesavelist)
                    print()

                    # 3 Create *.docx
                    if not glcombine:
                        exportfilename = os.path.join(downloadDirectory, f'Wiki - {wikiname}.docx')
                        WriteDocx(wikiDescription, exportfilename)
                    else:
                        wikiDescription = f'<h1>Сохранение: Wiki {wikiname}</h1>{wikiDescription}'
                        wikiCombineDescription = wikiCombineDescription + wikiDescription


        print(f'{Fore.CYAN}_________________________________________________________')


    # Combine one file
    if glcombine:
        exportfiletype = 'Issue'
        exporfilecontent = issueCombineDescription
        if gliswiki:
            exportfiletype = 'Wiki'
            exporfilecontent = wikiCombineDescription

        exportfilename = os.path.join(downloadDirectory, f'{exportfiletype}Combine.docx')
        WriteDocx(exporfilecontent, exportfilename)

    # Delete Img file
    if not glsaveimg:
        for file in imagesavelist:
            if os.path.isfile(file):
                os.remove(file)

    print(f'{Fore.CYAN}Process completed, press Space...')
    keyboard.wait("space")


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
