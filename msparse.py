#!/usr/bin/env python

from bs4 import BeautifulSoup
from collections import OrderedDict as OD
from os import path,mkdir
import sys
import json
import requests

#url = "https://technet.microsoft.com/en-us/library/security/ms16-027.aspx"
url_ms = "https://technet.microsoft.com/en-us/library/security/"
#r = requests.get(url)

# TODO this is nasty, refactor
def webfetch(report, subdir='', url=url_ms, ret=False):
    FILE = path.join(subdir, report + '.html')
    req_obj = requests.get(url+report)
    response= req_obj.text
    if not ret:
      ofh = open(FILE, 'wb')
      ofh.write(response.encode('utf8'))
      ofh.close()
      return
    return response


# Eventually....
# parse.py apr
#   fetch the montly summaries (ms16-apr) and parse the Executive Summary and Exploitability Index sections
#   using the parsed data, fetch all the individual bulletins (ms16-050 et al) that are mentioned and parse them
# Cache everything to disk
#   summaries/
#   bulletins/

# For now, we can feed cached files to this monstrosity

#   Need:
# The executive summary, workarounds, mitigations, kbs, cves, list of vuln products.
# Bulletin ; Description ; MS Rating ; BNS Rating ; ABM Rating ; Affected Products ; Potential Impact ; Details To Support BNS Rating
#  ExSum   ; ExSum(bold) ; ExSum*2   ;   ????     ;   ?????    ;   ExSum           ;  ExSum           ; *Executive Summary
#             MS###{                                               MS###: all lstd ;                  ; *Threat Attack Vectors
#            Upd.Replace                                                                              ; *Mitigating Factors   
#            KB Link                                                                                  ; *ABM Rating (Other): ALWAYS N/A
#            CVE List}                                                                                ; *BNS Other: ALWAYS N/A

def getText(node):
  '''Takes the text 'twixt the tags. Fix some unicode while we're at it.'''
  theText = node.getText().strip().replace(u'\xa0',' ').replace(u'\u2019',"'").replace('\n\n','|')
  try:
    theText.index('|')
    retval = theText.split('|')
  except ValueError:
    retval = theText
  return retval

def handleExecSummaries(table):
  '''The first table in the monthly summaries. Pretty standard.'''
  tdata = []
  rows = table.find_all('tr')
  headings = []
  for col in rows[0].find_all('td'):
    headings.append(getText(col))
  for row in rows[1:]:
    entry = OD()
    cols = row.find_all('td')
    for idx in xrange(0,len(cols)):
      entry[headings[idx]] = getText(cols[idx])
    tdata.append(entry)
  return tdata

def handleExploitIndex(table):
  '''The second table in the monthly summaries. A little funky but not bad.'''
  #tdata = []
  tdata = OD()
  rows = table.find_all('tr')
  headings = []
  for col in rows[0].find_all('td'):
    headings.append(getText(col))
  entry = None
  for row in rows[1:]:
    cols = row.find_all('td')
    if len(cols) is 1:
      if entry is not None: tdata[bulletin] = OD([("Bulletin",bulletin), ("Description",desc),("Details",entries)])
      bulletin,desc = getText(cols[0]).split(": ",1)
      entries = []
      entry = OD()
    else:
      for idx in xrange(0,len(cols)):
        entry[headings[idx]] = getText(cols[idx])
      entries.append(entry)
      entry = OD()
  return tdata


# Here are the annoyingly variable handlers for the sections found in the bulletins (i.e. ms16-050)
# Each handler may need two or more flavors implemented for cases when the section ctains simple text or more elaborate info (tables/bulleted lists/etc)

# TODO fill all these in :/
def handleBExecSumm(div, name=None):
  '''Executive Summary'''
  data = div.getText().strip()
  return 'Executive Summary', data

def handleBVulnInfo(div, name=None):
  '''Vulnerability Information'''
  # This one varies widely too... somtimes simple text, sometimes complex tables with additional subsections
  #data = div.getText().strip()
  data = list()
  for span in div.find_all('span'):
    for p in span.find_all('p'):
      data.append(p.getText().strip())
  return 'Vulnerability Information', data

def handleBAffSoft(div, name=None):
  '''Affected Software'''
  # another table :/
  data = None
  return 'Affected Software', data

def handleBFAQ(div, name=None):
  '''Frequently Asked Questions'''
  # Mostly text, but sometimes tables...
  # see ms16-001 and 050
  data = None
  return 'FAQ', data

def handleBMitFact(div, name=None):
  '''Mitigating Factors'''
  data = []
  for bullet in div.find_all('li'):
    getText(bullet)
    data.append(getText(bullet))
  return 'Mitigating Factors', data

def handleBWorkArnd(div, name=None):
  '''Workarounds'''
  # 'li' holds a 'strong' title and 'li' steps
  data = None
  return 'Workarounds', data

def handleBSecUpDepl(div, name=None):
  '''Security Update Deployment'''
  ps = div.find_all('p')
  data = []
  for p in ps:
    data.append(getText(p))
  return 'Security Update Deployment', data

def handleBAcks(div, name=None):
  '''Acknowledgements'''
  # To fill this info we need to link to another page with a table showing the MS#, title, CVEs, and free-form acknowledgement text
  data = getText(div.find('p'))
  return 'Acknowledgements', data

def handleBDsclmr(div, name=None):
  '''Disclaimer'''
  data = getText(div.find('p'))
  return 'Disclaimer', data

def handleBRevs(div, name=None):
  '''Revisions'''
  lis = div.find_all('li')
  data = []  
  for li in lis:
    data.append(getText(li))
  return 'Revisions', data

def handleBASVSR(div, name=None):
  '''Affected Software and Vulnerability Severity Ratings'''
  data = None
  return 'Affected Software and Vulnerability Severity Ratings', data

def handleBSRVI(div, name=None):
  '''Severity Ratings and Vulnerability Identifiers'''
  data = None
  return 'Severity Ratings and Vulnerability Identifiers', data

def handleBNotFound(div, name):
  '''If we come across a TOC entry we don't know about, call this'''
  print >>sys.stderr, "Unhandled TOC entry: %s" % name
  data = None
  return 'TOC ENTRY HANDLER NOT FOUND', data

def handleBulletin(msid):
  '''Top-level handler for the MS-### bulletins.
  Check ./bulletins/ for the cached file or fetch it fresh from MS
  Also pull ISC Rating from SANS'''

  FILE = path.join('bulletins', msid + '.html')
  if path.isfile(FILE) and (path.getsize(FILE) > 0):
    print >>sys.stderr, "Requested bulletin already cached: %s" % FILE
  else:
    webfetch(msid,'bulletins')
  
  sans_url = "https://isc.sans.edu/api/getmspatch/%s?json" % msid
  sans_json = webfetch('', url=sans_url, ret=True)
  sans_json = json.loads(sans_json)

  html = open(FILE, "r").read()
  soup = BeautifulSoup(html)

  # Each TOC entry gets a handler. MS is a bit inconsisent with the names so some handlers have multiple entries
  # The handlers should return a consistent name and their parsed data
  sectionHandlers = {
     "Executive Summary":           handleBExecSumm,
     "Vulnerability Information":   handleBVulnInfo,
     "Affected Software":           handleBAffSoft,
     "Frequently Asked Questions":  handleBFAQ,
     "Update FAQ":                  handleBFAQ,
     "Update FAQs":                 handleBFAQ,
     "Mitigating Factors":          handleBMitFact,
     "Workarounds":                 handleBWorkArnd,
     "Security Update Deployment":  handleBSecUpDepl,
     "Affected Software and Vulnerability Severity Ratings": handleBASVSR,
     "Severity Ratings and Vulnerability Identifiers":       handleBSRVI,
     "Acknowledgments":             handleBAcks,
     "Acknowledgements":            handleBAcks,
     "Disclaimer":                  handleBDsclmr,
     "Revisions":                   handleBRevs,
     }

  bulletinData = OD()

  # Iterate through the table of contents to get a pointer to where the section starts
  # Pass that section to its associated handler and return a dictionary{} with the section name and its parsed contents
  toc = soup.find("div",{"class":"sidebar_toc"})
  toc_entries = toc.find_all('li')
  for entry in toc_entries:
    entry_name = entry.getText().strip()
    entry_id = entry.a['href'].split("#",1)[1]
    span = soup.find(id=entry_id)
    section = span.findNext('div')
    name, data = sectionHandlers.get(entry_name, handleBNotFound)(section, name=entry_name)
    bulletinData[name] = data
  bulletinData["SANS"] = sans_json
  return bulletinData

if __name__ == '__main__':
  MTH = sys.argv[1] # 3-letter month: jan feb mar apr may jun 
  YEAR = '16'

  REPORT = 'ms'+YEAR+'-'+MTH
  FILE = REPORT+'.html'

  if path.isfile(FILE) and (path.getsize(FILE) > 0):
    print >>sys.stderr, "Requested month already cached: %s" % FILE
  else:
    webfetch(REPORT)

  if not path.isdir('bulletins'):
    mkdir('bulletins')

  html = open(FILE,'r').read()
  soup = BeautifulSoup(html)
  tables = soup.find_all('table')
  data = OD()

  data["Executive Summaries" ] = handleExecSummaries(tables[0])
  data["Exploitability Index"] = handleExploitIndex(tables[1])
  data["Bulletins"] = OD()

  for entry in data["Executive Summaries"]:
    msid = entry["Bulletin ID"]
    data["Bulletins"][msid] = handleBulletin(msid)
  print json.dumps(data)
  

# Tmporarily disabl parsing the montly summaries while working on the bulletins
def Ignore():
  msid = sys.argv[1]
  # save a bulletin in a bulletins/ subdirectory.
  #   bulletins/ms16-050.html
  # ./parse.py ms16-050 > ms16-050.json
  data = handleBulletin(msid)
  #print data
  print json.dumps(data)


  # change this to take the month as a param and fetch directly, and optional file input
  html = open(sys.argv[1], "r").read()
  soup = BeautifulSoup(html)
  # may need a better way to find these tables if the pattern changes...
  tables = soup.find_all('table')
  data = OD()

  data["Executive Summaries" ] = handleExecSummaries(tables[0])
  data["Exploitability Index"] = handleExploitIndex(tables[1])
  print json.dumps(data)

# we don't care about these tables, just the bulletin detail tables yay
#dontdo ... the remaining 5 funky tables :/
# Windows Operating Systems and Components (Table 1 of 2)
# Windows Operating Systems and Components (Table 2 of 2)
# Microsoft Office Suites and Software 
# Microsoft Office Services and Web Apps
# Microsoft Server Software





