#! python3
## PZLetterWriter: spits out all the required notice letters, all at once.

import datetime, docx, os, subprocess
print('Where should I save these?')
newdir=str(input())
os.chdir(newdir)
print('What is the P&Z meeting date? (Month DD, YYYY)')
PZDate=input()
print('What is the application number?')
appNum=str(input())
print('What is the project name?')
projName=str(input())
print('How many lots?')
totalLots=str(input())
print('How many residential lots?')
resLots=str(input())
print('How many common lots?')
comLots=str(input())
print('What is the site acreage?')
siteAcres=str(input())
print('What is the applicant\'s name?')
applicant=str(input())
print('What is the applicant\'s street address?')
applicantStreet=str(input())
print('What is the applicant\'s city?')
applicantCity=str(input())
print('What is the applicant\'s state?')
applicantState=str(input())
print('What is the applicant\'s zip code?')
applicantZip=str(input())
print('What is the applicant\'s phone number?')
applicantPhone=str(input())
print('What is the applicant\'s email?')
applicantEmail=str(input())
print('Who is the representative?')
repName=str(input())
if repName!=applicant:
    print('What is the representative\'s street address?')
    repStreet=str(input())
    print('What is the representative\'s city?')
    repCity=str(input())
    print('What is the representative\'s state?')
    repState=str(input())
    print('What is the representative\'s zip code?')
    repZip=str(input())
    print('What is the representative\'s phone number?')
    repPhone=str(input())
    print('What is the representative\'s email?')
    repEmail=str(input())
else:
    repStreet=applicantStreet
    repCity=applicantCity
    repState=applicantState
    repZip=applicantZip
    repPhone=applicantPhone
    repEmail=applicantEmail
print('Fill in the blank: '+appNum+' - '+projName+' - ' +applicant+': '+applicant+ ' represented by '+repName+'is requesting ____'+' for '+projName+', a '+totalLots+'-lot residential subdivision ('+resLots+' residential, '+comLots+' common).')
subjectBlank=input()
subject=applicant+ ' represented by '+repName+' is requesting '+subjectBlank+' for '+projName+', a '+totalLots+'-lot residential subdivision ('+resLots+' residential, '+comLots+' common).'
print('Fill in the blank: '+'The '+siteAcres+'-acre site is generally located ____')
locationBlank=input()
location='The '+siteAcres+'-acre site is generally located '+locationBlank
print('Who is the staff contact?')
staff=str(input())
print('What is the staff contact\'s title?')
staffTitle=str(input())
print('What is the staff contact\'s email?')
staffEmail=str(input())
print('When should this be published? (Month DD, YYYY)')
pubDate=str(input())
today=datetime.datetime.now()
##ENTITY LETTER
applicantAdd=(applicant+'\n'+applicantStreet+'\n'+applicantCity+', '+applicantState+' '+applicantZip)
repAdd=(repName+'\n'+repStreet+'\n'+repCity+', '+repState+' '+repZip)
ent=docx.Document('K:\\Planning Dept\\Base Documents\\entBase.docx')
transTable=ent.tables[0]
transDate=transTable.cell(0,1)
transDate.text=today.strftime('%B %d, %Y')
meetTable=ent.tables[1]
meetDate=meetTable.cell(0,1)
meetDate.text=PZDate
appTable=ent.tables[2]
appNumCell=appTable.cell(0,1)
appNumCell.text=appNum
projTable=ent.tables[3]
projDesc=projTable.cell(0,1)
projDesc.text=projName
addTable=ent.tables[4]
appAddCell=addTable.cell(1,0)
appAddCell.text=applicantAdd+'\n'+applicantPhone+'\n'+'Email: '+applicantEmail
repAddCell=addTable.cell(1,1)
repAddCell.text=repAdd+'\n'+repPhone+'\n'+'Email: '+repEmail
subTable=ent.tables[5]
subCell=subTable.cell(0,0)
subCell.text='SUBJECT: '+appNum+' - '+projName+' - ' +applicant+': ' +subject+' '+location
staffTable=ent.tables[6]
staffNameCell=staffTable.cell(0,1)
staffNameCell.text=staff+', '+staffTitle
staffEmailCell=staffTable.cell(0,2)
staffEmailCell.text=staffEmail

ent.save(projName+' ent.docx')
print('ENT LETTER SAVED')

##300 LETTER
threeHundred=docx.Document('K:\\Planning Dept\\Base Documents\\300Base.docx')
datePara=threeHundred.paragraphs[7]
datePara.add_run(today.strftime('%B %d, %Y'))
applicantPara=threeHundred.paragraphs[9]
applicantPara.add_run(applicant)
subLocPara=threeHundred.paragraphs[11]
subLocPara.add_run(appNum+' - '+projName+' - ' +applicant+': ').underline=True
subLocPara.add_run(' '+subject+' '+location)
meetPara=threeHundred.paragraphs[22]
meetPara.add_run('\t'+'\t')
meetPara.add_run(PZDate).underline=True
meetPara.add_run('\t'+'\t+TIME: 6:00 p.m.')
staffPara=threeHundred.paragraphs[31]
staffPara.add_run(staff)
titlePara=threeHundred.paragraphs[32]
titlePara.add_run(staffTitle)

threeHundred.save(projName+' 300.docx')
print('300 LETTER SAVED')

##REP LETTER
rep=docx.Document('K:\\Planning Dept\\Base Documents\\repBase.docx')
datePara=rep.paragraphs[7]
datePara.add_run(today.strftime('%B %d, %Y'))
repAddPara=rep.paragraphs[9]
repAddPara.add_run(repAdd)
meetPara=rep.paragraphs[12]
meetPara.add_run(PZDate)
subLocPara=rep.paragraphs[14]
subLocPara.add_run(appNum+' - '+projName+' - ' +applicant+': ').underline=True
subLocPara.add_run(' '+subject+' '+location)
staffPara=rep.paragraphs[31]
staffPara.add_run(staff)
titlePara=rep.paragraphs[32]
titlePara.add_run(staffTitle)

rep.save(projName+' rep.docx')
print('REP LETTER SAVED')

##PUB LETTER
pub=docx.Document('K:\\Planning Dept\\Base Documents\\pubBase.docx')
datePara=pub.paragraphs[9]
datePara.add_run(today.strftime('%B %d, %Y'))
pubDatePara=pub.paragraphs[11]
pubDatePara.add_run(pubDate)
meetPara=pub.paragraphs[16]
meetPara.add_run('Legal notice is hereby given that the EAGLE PLANNING AND ZONING COMMISSION will hold a public hearing '+PZDate+', at 6:00 P.M. at Eagle City Hall to consider the following:')
applicationPara=pub.paragraphs[18]
applicationPara.add_run(appNum)
applicantPara=pub.paragraphs[20]
applicantPara.add_run(applicant)
subjectPara=pub.paragraphs[22]
subjectPara.add_run(appNum+' - '+projName+' - ' +applicant+': ').underline=True
subjectPara.add_run(subject)
locationPara=pub.paragraphs[24]
locationPara.add_run(location)

pub.save(projName+' pub.docx')
print('PUB LETTER SAVED')
print('Letters saved in requested location!')
##TODO: Pop Open Results

input('Press ENTER to exit.')
