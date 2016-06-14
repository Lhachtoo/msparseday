#!/usr/bin/env python

#from docx import Document
import sys
import dominate
from dominate.tags import *

import json
from collections import OrderedDict as OD


#apr = json.load(open("201604.json"), object_pairs_hook=OrderedDict)

#In [7]: apr.keys()
#Out[7]: [u'Executive Summaries', u'Exploitability Index']

#doc = dominate.document(title='Patch Tuesday HTML Test')

def docheader(idoc):
  with idoc.head:
#    link(rel='stylesheet', href='style.css')
#    #script(type='text/javascript', src='script.js')
#  style("body{font-family:Helvetica;font-size:small}")
    style("""
@font-face
        {font-family:Wingdings;
        panose-1:5 0 0 0 0 0 0 0 0 0;}
@font-face
        {font-family:Wingdings;
        panose-1:5 0 0 0 0 0 0 0 0 0;}
@font-face
        {font-family:Cambria;
        panose-1:2 4 5 3 5 4 6 3 2 4;}
@font-face
        {font-family:Calibri;
        panose-1:2 15 5 2 2 2 4 3 2 4;}
@font-face
        {font-family:Tahoma;
        panose-1:2 11 6 4 3 5 4 4 2 4;}
@font-face
        {font-family:Verdana;
        panose-1:2 11 6 4 3 5 4 4 2 4;}
@font-face
        {font-family:Consolas;
        panose-1:2 11 6 9 2 2 4 3 2 4;}
@font-face
        {font-family:"Segoe UI";
        panose-1:2 11 5 2 4 2 4 2 2 3;}
@font-face
        {font-family:"Lucida Sans Unicode";
        panose-1:2 11 6 2 3 5 4 2 2 4;}
 /* Style Definitions */

table, th, td
        {border: 1px solid black}

table.MsoNormalTable
        {border:1;
        cellspacing:0;cellpadding:0;align:left;
        width:1267;style:width:13.2in;border-collapse:collapse;border:none;
        margin-left:6.75pt;margin-right:6.75pt}

 p.MsoNormal, li.MsoNormal, div.MsoNormal
        {margin:0in;
        margin-bottom:.0001pt;
        font-size:12.0pt;
        font-family:"Times New Roman","serif";}

tr.MsoHeaderRow
        {width:1.45in;border:solid windowtext 1.0pt;background:blue;color:yellow;
        padding:0in 5.4pt 0in 5.4pt;height:22.0pt;}
h1
        {mso-style-link:"Heading 1 Char";
        margin-top:12.0pt;
        margin-right:0in;
        margin-bottom:3.0pt;
        margin-left:0in;
        page-break-after:avoid;
        font-size:16.0pt;
        font-family:"Arial","sans-serif";
        font-weight:bold;}
h2
        {mso-style-link:"Heading 2 Char";
        margin-top:12.0pt;
        margin-right:0in;
        margin-bottom:3.0pt;
        margin-left:0in;
        page-break-after:avoid;
        font-size:14.0pt;
        font-family:"Arial","sans-serif";
        font-weight:bold;
        font-style:italic;}
h3
        {mso-style-link:"Heading 3 Char";
        margin-top:12.0pt;
        margin-right:0in;
        margin-bottom:3.0pt;
        margin-left:0in;
        page-break-after:avoid;
        font-size:13.0pt;
        font-family:"Cambria","serif";
        font-weight:bold;}
h4
        {mso-style-link:"Heading 4 Char";
        margin-top:12.0pt;
        margin-right:0in;
        margin-bottom:3.0pt;
        margin-left:0in;
        page-break-after:avoid;
        font-size:14.0pt;
        font-family:"Calibri","sans-serif";
        font-weight:bold;}
p.MsoFootnoteText, li.MsoFootnoteText, div.MsoFootnoteText
        {mso-style-link:"Footnote Text Char";
        margin:0in;
        margin-bottom:.0001pt;
        font-size:10.0pt;
        font-family:"Times New Roman","serif";}
p.MsoHeader, li.MsoHeader, div.MsoHeader
        {mso-style-link:"Header Char";
        margin:0in;
        margin-bottom:.0001pt;
        font-size:12.0pt;
        font-family:"Times New Roman","serif";}
p.MsoFooter, li.MsoFooter, div.MsoFooter
        {mso-style-link:"Footer Char";
        margin:0in;
        margin-bottom:.0001pt;
        font-size:12.0pt;
        font-family:"Times New Roman","serif";}
span.MsoFootnoteReference
        {font-family:"Times New Roman","serif";
        vertical-align:super;}
p.MsoSubtitle, li.MsoSubtitle, div.MsoSubtitle
        {mso-style-link:"Subtitle Char";
        margin-top:0in;
        margin-right:0in;
        margin-bottom:3.0pt;
        margin-left:0in;
        text-align:center;
        font-size:12.0pt;
        font-family:"Cambria","serif";}
a:link, span.MsoHyperlink
        {font-family:"Times New Roman","serif";
        color:blue;
        text-decoration:underline;}
a:visited, span.MsoHyperlinkFollowed
        {font-family:"Times New Roman","serif";
        color:purple;
        text-decoration:underline;}
strong
        {font-family:"Times New Roman","serif";}
p.MsoDocumentMap, li.MsoDocumentMap, div.MsoDocumentMap
        {mso-style-link:"Document Map Char";
        margin:0in;
        margin-bottom:.0001pt;
        font-size:8.0pt;
        font-family:"Tahoma","sans-serif";}
p.MsoPlainText, li.MsoPlainText, div.MsoPlainText
        {mso-style-link:"Plain Text Char";
        margin:0in;
        margin-bottom:.0001pt;
        font-size:10.5pt;
        font-family:Consolas;}
p
        {margin-right:0in;
        margin-left:0in;
        font-size:12.0pt;
        font-family:"Times New Roman","serif";}
p.MsoAcetate, li.MsoAcetate, div.MsoAcetate
        {mso-style-link:"Balloon Text Char";
        margin:0in;
        margin-bottom:.0001pt;
        font-size:8.0pt;
        font-family:"Tahoma","sans-serif";}
p.MsoListParagraph, li.MsoListParagraph, div.MsoListParagraph
        {margin-top:0in;
        margin-right:0in;
        margin-bottom:0in;
        margin-left:.5in;
        margin-bottom:.0001pt;
        font-size:12.0pt;
        font-family:"Times New Roman","serif";}
p.MsoListParagraphCxSpFirst, li.MsoListParagraphCxSpFirst, div.MsoListParagraphCxSpFirst
        {margin-top:0in;
        margin-right:0in;
        margin-bottom:0in;
        margin-left:.5in;
        margin-bottom:.0001pt;
        font-size:12.0pt;
        font-family:"Times New Roman","serif";}
p.MsoListParagraphCxSpMiddle, li.MsoListParagraphCxSpMiddle, div.MsoListParagraphCxSpMiddle
        {margin-top:0in;
        margin-right:0in;
        margin-bottom:0in;
        margin-left:.5in;
        margin-bottom:.0001pt;
        font-size:12.0pt;
        font-family:"Times New Roman","serif";}
p.MsoListParagraphCxSpLast, li.MsoListParagraphCxSpLast, div.MsoListParagraphCxSpLast
        {margin-top:0in;
        margin-right:0in;
        margin-bottom:0in;
        margin-left:.5in;
        margin-bottom:.0001pt;
        font-size:12.0pt;
        font-family:"Times New Roman","serif";}
span.MsoSubtleReference
        {font-variant:small-caps;
        color:#9B2D1F;
        text-decoration:underline;}
span.Heading1Char
        {mso-style-name:"Heading 1 Char";
        mso-style-link:"Heading 1";
        font-family:"Cambria","serif";
        font-weight:bold;}
span.Heading2Char
        {mso-style-name:"Heading 2 Char";
        mso-style-link:"Heading 2";
        font-family:"Cambria","serif";
        font-weight:bold;
        font-style:italic;}
span.Heading3Char
        {mso-style-name:"Heading 3 Char";
        mso-style-link:"Heading 3";
        font-family:"Cambria","serif";
        font-weight:bold;}
span.Heading4Char
        {mso-style-name:"Heading 4 Char";
        mso-style-link:"Heading 4";
        font-family:"Calibri","sans-serif";
        font-weight:bold;}
span.FootnoteTextChar
        {mso-style-name:"Footnote Text Char";
        mso-style-link:"Footnote Text";
        font-family:"Times New Roman","serif";}
span.BalloonTextChar
        {mso-style-name:"Balloon Text Char";
        mso-style-link:"Balloon Text";
        font-family:"Times New Roman","serif";}
p.lastincell, li.lastincell, div.lastincell
        {mso-style-name:lastincell;
        margin:0in;
        margin-bottom:.0001pt;
        line-height:130%;
        font-size:8.5pt;
        font-family:"Verdana","sans-serif";}
span.HeaderChar
        {mso-style-name:"Header Char";
        mso-style-link:Header;
        font-family:"Times New Roman","serif";}
span.FooterChar
        {mso-style-name:"Footer Char";
        mso-style-link:Footer;
        font-family:"Times New Roman","serif";}
span.PlainTextChar
        {mso-style-name:"Plain Text Char";
        mso-style-link:"Plain Text";
        font-family:Consolas;}
span.SubtitleChar
        {mso-style-name:"Subtitle Char";
        mso-style-link:Subtitle;
        font-family:"Cambria","serif";}
span.DocumentMapChar
        {mso-style-name:"Document Map Char";
        mso-style-link:"Document Map";
        font-family:"Tahoma","sans-serif";}
p.list6, li.list6, div.list6
        {mso-style-name:list6;
        margin-top:0in;
        margin-right:6.8pt;
        margin-bottom:0in;
        margin-left:6.8pt;
        margin-bottom:.0001pt;
        font-size:12.0pt;
        font-family:"Times New Roman","serif";}
span.apple-converted-space
        {mso-style-name:apple-converted-space;}
p.Default, li.Default, div.Default
        {mso-style-name:Default;
        margin:0in;
        margin-bottom:.0001pt;
        text-autospace:none;
        font-size:12.0pt;
        font-family:"Arial","sans-serif";
        color:black;}
span.CSSBodyChar
        {mso-style-name:"CSS Body Char";
        mso-style-link:"CSS Body";
        font-family:"Arial","sans-serif";
        color:#636363;}
p.CSSBody, li.CSSBody, div.CSSBody
        {mso-style-name:"CSS Body";
        mso-style-link:"CSS Body Char";
        margin:0in;
        margin-bottom:.0001pt;
        line-height:13.0pt;
        font-size:11.0pt;
        font-family:"Arial","sans-serif";
        color:#636363;}
 /* Page Definitions */
 @page WordSection1
        {size:14.0in 8.5in;
        margin:27.35pt 59.75pt 22.5pt 27.35pt;}
div.WordSection1
        {page:WordSection1;}
 /* List Definitions */
 ol
        {margin-bottom:0in;}
ul
        {margin-bottom:0in;}
""")

#   Need:
# The executive summary, workarounds, mitigations, kbs, cves, list of vuln products.
# Bulletin ; Description ; MS Rating ; BNS Rating ; ABM Rating ; Affected Products ; Potential Impact ; Details To Support BNS Rating
#  ExSum   ; ExSum(bold) ; ExSum*2   ;   ????     ;   ?????    ;   ExSum           ;  ExSum           ; *Executive Summary
#             MS###{                                               MS###: all lstd ;                  ; *Threat Attack Vectors
#            Upd.Replace                                                                              ; *Mitigating Factors   
#            KB Link                                                                                  ; *ABM Rating (Other): ALWAYS N/A
#            CVE List}                                                                                ; *BNS Other: ALWAYS N/A

def fixup_data(data):
  all_info = []

  # later sorting: foo = OrderedDict(sorted(foo.iteritems(), key=lambda x: x[1]['depth']))
  for entry in data["Executive Summaries"]:
    #line["Bulletin"] = entry["Bulletin ID"]
    line = OD()

    bid = entry["Bulletin ID"]
    affprod = entry["Affected Software"]
    msr,potimp = entry["Maximum Severity Ratingand Vulnerability Impact"].split(" ",1)
    descsumm = entry["Bulletin Title and Executive Summary"].split(")",1)
    desc = descsumm[0]
    summ = ""
    try:
      desc,summ = descsumm
      desc = desc + ')'
    except ValueError:
      pass

    kbnum = data["Bulletins"][bid]["Executive Summary"].rsplit(" ",1)[1].replace(".","")

    cves = list()
    try:
      for cventry in data["Exploitability Index"][bid]["Details"]:
        cves.append(cventry["CVE ID"]) 
    except KeyError:
      pass

    tav = ("FIXME_TAV",)
    abm = "N/A"
    bns = "N/A"

    try:
      miti = data["Bulletins"][bid]["Mitigating Factors"]
    except KeyError:
      miti = ("Microsoft has not identified any workarounds for this vulnerability",)

    line["Bulletin"] = bid
    line["Description"] = desc,kbnum,cves
    line["MS Rating"] = msr
    line["BNS Rating"] = "FIXME_BNS"
    line["ABM Rating"] = "FIXME_ABM"
    line["Affected Product(s)"] = affprod
    line["Potential Impact"] = potimp
    line["Details To Support BNS Rating"] = summ,tav,miti,abm,bns
    all_info.append(line)
  return all_info

BulletinTableHeaders = ("Bulletin", "Description", "MS Rating", "BNS Rating", "ABM Rating",
                        "Affected Product(s)", "Potential Impact", "Details To Support BNS Rating")

BulletinTableStyle = {
                 "Bulletin":"width:63.0pt;",
                 "Description":"width:1.95in;border-left:none;",
                 "MS Rating":"width:58.5pt;border-left:none;",
                 "BNS Rating":"width:.75in;border-left:none;",
                 "ABM Rating":"width:.5in;border-left:none;",
                 "Affected Product(s)":"width:120.6pt;border-left:none;",
                 "Potential Impact":"width:71.2pt;border-left:none;",
                 "Details To Support BNS Rating":"width:437.2pt;border-left:none;",
                 }

def mkmsotable(data):
  doc = dominate.document(title='Patch Tuesday HTML Test')
  #with doc.head:
  #  style("body{font-family:Helvetica;font-size:small}")
  docheader(doc)

  with doc.add(body()):
    h1('Test Header')
    p("Here's some text to explain this section")
    # Construct the main bulletin table
    with table(_class='MsoNormalTable').add(tbody()):
      lh = tr(_class='MsoHeaderRow')
      colw = len(BulletinTableHeaders)
      #header row
      for col in BulletinTableHeaders:
        lh += th(col, style=BulletinTableStyle[col])
      #content
      info = fixup_data(data)
      for row in info:
        lc = tr(_class='MsoContentRow')
        for col in BulletinTableHeaders:
          if col == "Description":
            # Hide bullets:: li or ul {list-style-type:none;}
            lc += td(p(strong(row[col][0])),
                     p(a("KB"+row[col][1],href="http://support.microsoft.com/kb/"+row[col][1])),
                     (li(a(cve,href="http://www.cve.mitre.org/cgi-bin/cvename.cgi?name="+cve)) for
                         cve in row[col][2]), style=BulletinTableStyle[col])
          elif col == "Details To Support BNS Rating":
            summ, tav, miti, abm, bns = row[col]
            lc += td(p(strong("Executive Summary:"), summ),
                     p(strong("Threat/Attack Vector(s):"), li(tav)),
                     p(strong("Mitigating Factor(s):"), li(miti)),
                     p(strong("ABM Rating (Other):"), li(abm)),
                     p(strong("BNS Other:"), li(bns)),
                    style=BulletinTableStyle[col])
          else:
            lc += td(p(row[col]), style=BulletinTableStyle[col])

  return doc

if __name__ == "__main__":
  data = json.load(open(sys.argv[1]), object_pairs_hook=OD)
  doc = mkmsotable(data)
  
  print doc

