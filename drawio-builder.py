import xml.etree.ElementTree as ET
import random
import string
import sys
import pandas as pd
import math

class drawioBuilder():

    def __init__(self):
        self.file = ET.Element('mxfile')
    
        for k,v in [
        ("host", "app.diagrams.net"),
        ("modified", "2023-09-05T07:54:04.190Z"),
        ("agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.0.0 Safari/537.36"),
        ("etag", self.get_random_string(8)),
        ("version", "21.7.2"),
        ("type", "device")]:
            self.file.set(k,v)
            
        d1 = ET.SubElement(self.file,'diagram')
        
        for k,v in [
        ("id", self.get_random_string(8)),
        ("name", "Page-1")
        ]:
            d1.set(k,v)
        
        graph = ET.SubElement(d1,"mxGraphModel")
        
        for k,v in [
        ("dx", "682"),
        ("dy", "680"),
        ("grid", "1"),
        ("gridSize", "10"),
        ("guides", "1"),
        ("tooltips", "1"),
        ("connect", "1"),
        ("arrows", "1"),
        ("fold", "1"),
        ("page", "1"),
        ("pageScale", "1"),
        ("pageWidth", "4000"),
        ("pageHeight", "4000"),
        ("math", "0"),
        ("shadow", "0"),
        ("extFonts", "Permanent Marker^https://fonts.googleapis.com/css?family=Permanent+Marker")
        ]:
            graph.set(k,v)
            
        root = ET.SubElement(graph,"root")
        mxC1 = ET.SubElement(root,"mxCell")
        mxC2 = ET.SubElement(root,"mxCell")
        
        mxC1.set("id","0")
        mxC2.set("id","1")
        mxC2.set("parent","0")

    def add_table(self, tableName, x, y, atts):

        ####
        # ADDING TABLE TO THE PAGE
        ####

        root = self.file[0][0][0]

        table = ET.SubElement(root,'mxCell')
    
        for k,v in [
        ("id", self.get_random_string(8) + "-1"),
        ("value", tableName),
        ("style", "shape=table;startSize=30;container=1;collapsible=1;childLayout=tableLayout;fixedRows=1;rowLines=0;fontStyle=1;align=center;resizeLast=1;html=1;"),
        ("parent", "1"),
        ("vertex", "1")]:
            table.set(k,v)
            
        geo1 = ET.SubElement(table,'mxGeometry')
        rec1 = ET.SubElement(geo1,'mxRectangle')
        
        for k,v in [
        ("x", str(x)),
        ("y", str(y)),
        ("width", "180"),
        ("height", str(len(atts)*30+30)),
        ("as", "geometry")
        ]:
            geo1.set(k,v)
            
        for k,v in [
        ("x", str(x)),
        ("y", str(y)),
        ("width", "180"),
        ("height", "90"),
        ("as", "geometry"),
        ("width", "70"),
        ("height", "30"),
        ("as", "alternateBounds")
        ]:
            rec1.set(k,v)

        table_id = table.attrib['id'][:len(table.attrib['id'])-2] # store table id


        ####
        # Now we add all attributes to the table
        ####

        for i,at in enumerate(atts): # atts = [(value,key)]

            print(at)

            ## Text box and respective geometry
            txtBox = ET.SubElement(root, 'mxCell')
            for k,v in [("id", table_id+"-"+str(2+3*i)),
                ("value", ""),
                ("style", "shape=tableRow;horizontal=0;collapsible=0;startSize=0;swimlaneHead=0;swimlaneBody=0;fillColor=none;dropTarget=0;points=[[0,0.5],[1,0.5]];portConstraint=eastwest;top=0;left=0;right=0;bottom=1;"),
                ("parent", table_id+"-1"),
                ("vertex", "1")]:
                txtBox.set(k,v)
            
            geo = ET.SubElement(txtBox,'mxGeometry')
            for k,v in [("y", str(30+30*i)),
                ("width", "180"),
                ("height", "30"),
                ("as", "geometry")]:
                geo.set(k,v)

            ## Adding key field. Is this PK, FK...? AND resprective geometry/rectangle
            keyFld = ET.SubElement(root, 'mxCell')
            for k,v in [ ("id", table_id+"-"+str(2+3*i+1)),
                ("value", at[1]),
                ("style", "shape=partialRectangle;connectable=0;fillColor=none;top=0;left=0;bottom=0;right=0;fontStyle=1;overflow=hidden;whiteSpace=wrap;html=1;"),
                ("parent", table_id+"-"+str(2+3*i)),
                ("vertex", "1")]:
                keyFld.set(k,v)
            
            ## Geometry
            geo2 = ET.SubElement(keyFld,'mxGeometry')
            for k,v in [("width", "30"),
                ("height", "30"),
                ("as", "geometry")]:
                geo2.set(k,v)

            ## Rectangle
            rect2 = ET.SubElement(geo2,'mxRectangle')
            for k,v in [("width", "30"),
                ("height", "30"),
                ("as", "alternateBounds")]:
                rect2.set(k,v)

            ## Adding attribute name... Quite repetative, should wrap into funtion...
            valFld = ET.SubElement(root, 'mxCell')
            for k,v in [("id", table_id+"-"+str(2+3*i+2)),
                ("value", at[0]),
                ("style","shape=partialRectangle;connectable=0;fillColor=none;top=0;left=0;bottom=0;right=0;align=left;spacingLeft=6;overflow=hidden;whiteSpace=wrap;html=1;"),
                ("parent", table_id+"-"+str(2+3*i)),
                ("vertex", "1")]:
                valFld.set(k,v)
            
            ## Geometry
            geo3 = ET.SubElement(valFld,'mxGeometry')
            for k,v in [("x", "30"),
                ("width", "150"),
                ("height", "30"),
                ("as", "geometry")]:
                geo3.set(k,v)

            ## Rectangle
            rect3 = ET.SubElement(geo3,'mxRectangle')
            for k,v in  [("width", "150"),
                ("height", "30"),
                ("as", "alternateBounds")]:
                rect3.set(k,v)


    def get_random_string(self,length):
        # choose from all lowercase letter
        letters = string.ascii_lowercase
        result_str = ''.join(random.choice(letters) for i in range(length))
        
        return result_str
    
if __name__ == '__main__':

    ##
    ## Read in the excel file AND get the static values
    ##

    df = pd.read_excel("dmAH.xlsx",sheet_name='Transactions')
    # Set the column names to the values in the first row
    df.columns = df.iloc[0]

    # Drop the first row (old column names) from the DataFrame
    df = df.iloc[1:]

    # Reset the index if needed
    df.reset_index(drop=True, inplace=True)

    condition = df["Grouping 1"] == "Static"

    static_df = df[condition]
    uVals = static_df['Object'].unique()

    valAts = {}

    for v in uVals:
        valAts[v] = []
        tupes = list(zip(static_df[static_df['Object']==v]['Data Point'],static_df[static_df['Object']==v]['Dependency']))
        valAts[v] = tupes
    
    alts_refined = {} # Dictionary to store the vlues without nan vals

    for keys in valAts.keys():
        ## Remove nan values and replace with empty string
        to_remove = []
        for i,k in enumerate(valAts[keys]):
            if type(k[0]) is float and math.isnan(k[0]):
                to_remove.append(valAts[keys][i])
                
        alts_refined[keys] = [item for item in valAts[keys] if item not in to_remove]

        ## Remove nan values and replace with empty string
        for i,k in enumerate(alts_refined[keys]):
            if pd.isna(k[1]):
                alts_refined[keys][i] = (k[0],"")

    print(alts_refined)

    dio = drawioBuilder()

    # for i,k in enumerate(alts_refined.keys()):
    #     dio.add_table(k, 100+1000*(i%5), 100+1000*(i//5), alts_refined[k])

    dio.add_table("Table 1", 100, 100, [("Attribute 1","PK"),("Attribute 2","FK")])

    a = dio.file

    with open('tst.drawio','w') as f:
        sys.stdout = f
        ET.dump(a)
        
    sys.stdout = sys.__stdout__