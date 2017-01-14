# Author: John W. Haney, GISP
# Date: December 2015

# Description: Python script that utilizes arcpy to create a project file gdb, loop through all project .xlsx files and sheets (tables)
# in the given path, and convert all relevant tables to feature classes in the new project gdb.

import arcpy, os, time

# Allow for over-writing files
arcpy.env.overwriteOutput = True

# User inputs path to files and a variable is created for path
wspc = raw_input("Enter path to project excel file(s): ")

# Set environment workspace to path entered by user input above
arcpy.env.workspace = wspc

# Create variable that lists xlsx documents in workspace
tables = arcpy.ListFiles("*.xlsx")

# Create project file geodatabase variable
gdb = raw_input("Enter name of project .gdb w/ file extension: ")

# Create time variable
start_time = time.clock()

# Set spatial reference variable to use later
spref = arcpy.SpatialReference(4269)

# Create file geodatabase within specified workspace
arcpy.CreateFileGDB_management(arcpy.env.workspace, gdb)
print "gdb created successfully...\n"

# Print list of all xlsx excel documents within specified workspace
print "list of .xlsx files in {}...\n{}\n".format(wspc, tables)

try:
    print "in progress...\n"
    for xlsx in tables:
        arcpy.env.workspace = os.path.join(wspc, xlsx)  # Loop through every xlsx in path and set as workspace
        print "working in " + arcpy.env.workspace + "\n"   # Test
        sht_list = arcpy.ListTables()   # Create variable for all sheets (tables) inside xlsx
        print "list of tables from current .xlsx workspace...\n{}\n".format(sht_list)
        for sht in sht_list:  
            if "#" in sht:      # Skip over tables with invalid characters in name (usually caused by excel filters)
                continue
            else:
                tbl = sht.replace("'", "").replace("$", "")   # Create a variable with no special characters
                fc = "fc" + tbl    
                temp = "temp" + tbl    # Create temporary event layer
                print "there are {} rows in table {}\n".format(arcpy.GetCount_management(sht), tbl)
                arcpy.MakeXYEventLayer_management(sht, "Lon", "Lat", temp, spref)   # Make temporary XY Events layer
                print "creating {} featureclass...\n".format(fc)
                arcpy.FeatureClassToFeatureClass_conversion(temp, os.path.join(wspc, gdb), fc)   # Make Featureclass from XY Events
                print "{} successfully created.\n".format(fc)
    print "PROGRAM COMPLETED SUCCESSFULLY IN"
    print(" --- %s SECONDS ---" % (time.clock() - start_time))
except:
    print "Error!!!\n" + arcpy.GetMessages()    # Print any ArcPy specific error messages
