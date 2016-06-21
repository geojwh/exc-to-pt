# Author: John W. Haney, GISP
# Date: December 2015

# Description: Python script that utilizes arcpy to create a project file gdb, loop through all project .xlsx files and sheets (tables)
# in the given path, and convert all relevant tables to feature classes in the new project gdb.

# Preparation:
# (1) Only tables (sheets) with both "Lat" and "Lon" field names can be included, otherwise the program will print an error.
# (2) Insure there are no spaces or strange characters in the sheet names of each excel file. Give each sheet a descriptive
#     name such as "Crash2015".

# Instructions to run program:
# (1) Open a command window in the same directory that this script is saved - hold shift and right-click anywhere in the directory, choose
#     'open command window here'.
# (2) Type 'python exc-to-pt.py' and Press Enter. - this only works if you have already added python to the path in the system environment variables
#	  (Windows). If you have not done this, you will need to do so or use an IDE to run this script, such as IDLE or PyCharm, etc.
# (3) When prompted, enter the path to where your project xlsx files are saved. Press Enter.
# (4) When prompted, enter the name of your project file geodatabase, including the .gdb file extension. Press Enter.
# (5) If the program is running and completes successfully, you should see several print statements as the program is running, and then
#     a print statement that reads 'PROGRAM COMPLETED SUCCESSFULLY'. Navigate to your project path and view the gdb and featureclasses to
#     make sure everything is good to go. !Voila!

# *** Additional customization to this script can be completed in the future in order to include joining 'null' lat/lon values in DPS citation datasets
#     to a TXDOT mile marker file (method developed to enhance location information for null data points). That data will then be merged back to the original
#	  featureclass to compile a more complete dataset for project use.

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
