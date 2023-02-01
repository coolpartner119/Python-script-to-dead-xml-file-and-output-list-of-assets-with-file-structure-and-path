import os
import gzip
import xml.etree.ElementTree as ET
import csv
import pandas as pd
import openpyxl
from string import ascii_uppercase
INPUT_FILE = input("File Name: ")
OUTPUT_FILE = INPUT_FILE  + "output.xml"
ITEMS = []

myRoot = 0;
def file_execution():
  new_file = f"{INPUT_FILE}.gz"
  if(os.path.exists(f"{INPUT_FILE}.prproj")):
    os.rename(f"{INPUT_FILE}.prproj", new_file)

  f = gzip.open(new_file, 'r')
  file_content = f.read()
  file_content = file_content.decode('utf-8')
  f_out = open(OUTPUT_FILE, 'w+')
  f_out.write(file_content)
  f.close()
  f_out.close()

def parse_xml():
  global myRoot
  myTree =  ET.parse(OUTPUT_FILE)
  myRoot = myTree.getroot()
  bins = myRoot.findall("BinProjectItem")
  clips = myRoot.findall("ClipProjectItem")
  items = myRoot.find("RootProjectItem").find("ProjectItemContainer").find("Items").findall("Item")
  for item in items:
    objectId = item.attrib["ObjectURef"]
    specify_item(bins,clips,objectId,"Root/")


def specify_item(bins,clips,objectId,treePath):
  type = ""
  specific_item = 0;
  for bin in bins:
    if(bin.attrib["ObjectUID"] == objectId):
      type = "BIN"
      specific_item = bin
      break
  for clip in clips:
    if(clip.attrib["ObjectUID"] == objectId):
      type = "CLIP"
      specific_item = clip
      break
  if(type == "CLIP"):
    clip_name = parse_name(specific_item)
    clip_tree = treePath + clip_name
    clip_path = get_path(clip_name)
    clip_id = objectId
    
    clip = {"treePath": clip_tree, "name": clip_name, "type": "CLIP", "path": clip_path,"nodeId" : clip_id}
    ITEMS.append(clip)
  elif(type == "BIN"):
    bin_name = parse_name(specific_item)
    bin_tree = treePath + bin_name
    bin = {"treePath": bin_tree, "name": bin_name, "type": "BIN", "path": "","nodeId" : ""}
    ITEMS.append(bin)
    
    itemsList = specific_item.find("ProjectItemContainer").findall("Items")
    if(len(itemsList) > 0):
      newTreePath = bin_tree + "/"
      for items in itemsList:
        for item in items:
        
          objectId = item.attrib["ObjectURef"]
          specify_item(bins,clips,objectId,newTreePath)
        
    
def parse_name(item):
  return item.find("ProjectItem").find("Name").text

def get_path(objectName):
  medias = myRoot.findall("Media")
  for media in medias:
    path = media.find("RelativePath").text
    if(path.__contains__(objectName)):
      new_path = path.split("../")[-1].split("..\\")[-1]
      
      return new_path
    
  return False

def output():
  header = ["ProjectItem.treePath","ProjectItem","ItemType","Path to media file","ProjectItem.nodeId"]
  outputFile = f"{INPUT_FILE}_result.csv"
  f = open(outputFile,"w")
  writer = csv.writer(f)
  writer.writerow(header)
  for item in ITEMS:
    rowItem = [item["treePath"],item["name"],item["type"],item["path"],item["nodeId"]]
    writer.writerow(rowItem)
  f.close()
  read_file = pd.read_csv(outputFile)
  read_file.to_excel(f"{INPUT_FILE}_output.xlsx",index = None, header=True)

def adjust_cells():
  newFile = f"{INPUT_FILE}_output.xlsx"

  wb = openpyxl.load_workbook(filename = newFile)        
  worksheet = wb.active

  for column in ascii_uppercase:
    if (column=='A'):
        worksheet.column_dimensions[column].width = 80
    elif (column=='B'):
        worksheet.column_dimensions[column].width = 50           
    elif (column=='C'):
        worksheet.column_dimensions[column].width = 10    
    elif (column=='D'):
        worksheet.column_dimensions[column].width = 80  
    elif (column=='E'):
        worksheet.column_dimensions[column].width = 60          
    else:
        worksheet.column_dimensions[column].width = 50

  wb.save(newFile)

def main():
  file_execution()
  parse_xml()
  output()
  adjust_cells()
  print("Files have been generated successfully ! :)")

if __name__ == "__main__":
  main()