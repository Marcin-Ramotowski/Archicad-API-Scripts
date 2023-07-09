# Archicad-API-Scripts

Scripts to automate some operations in Archicad.

# Table of Contents

* ![General Info](#general-information)
* ![Technologies Used](#technologies-used)
* ![Features](#features)
* ![Setup](#setup)
* ![Usages](#usages)
* ![Project Status](#project-status)
* ![Acknowledgements](#acknowledgements)
* ![Contact](#contact)

# General Information

A set of Python scripts that connect to an open building project
in Archicad using the Archicad Automation API,
retrieve information from it,
and make changes to model elements.

# Technologies Used

Packages used for this project are listed in *requirements.txt*


# Features

List the ready scripts here:

Scripts to display information on console:
  - All properties - It displays names of all properties available in the project.
  - Check Properties - It displays values of all selected properties for all elements of the specified type.
  - Navigator Elements - It displays a tree showing the available elements in the navigator and their hierarchy.
  - Project Elements - It subtotals the elements of each type in the project and prints the results on the screen.
  - Property guid - It gets guid of selected property and displays on the screen.
  - Repeating IDs - It displays id values that occur more than once in the project.
  - Sorting - It sorts elements of a given type according to the first specified value in ascending or descending order.

Scripts saving data into new files:
  - Excel exporter - It creates a new .xlsx file and places in it the values of the indicated properties from the elements     highlighted in the opened Archicad's window.
  - Floor space distribution- It creates plots describing the share of individual zones in the area of the entire floor.
  - Room report - It collects data on the zone representing the room in question and inserts it into the indicated file at the indicated locations.
  - Table of furniture - It creates and saves worksheet consists of a list of objects in the project, broken down by the area in which they are located. It displays also information about them and calculating total and partial sums of prices for all of assigned objects.
  - Zone allocation - It assigns objects to the zones within which they are located based on data from the project and the exported IFC file.

Scripts that set properties of project elements:
  - Excel importer - It retrieves data from the specified .xlsx file and sets property values in the project according to that data.
  - IDs assignment - It assigns unique ids for objects according to Archicad's default ID allocation rules. Moreover to objects with the same dimensions assign the id of the first of these twin elements. Great for cleaning up the mess of item IDs.
  - Shared Id for walls - It assigns the same id for these walls which have the same construction composite.
  - Shared Id - It assigns to objects with the same dimensions the id of the first of these twin elements.
  - Zones numbering - It numbers the zones by assigning them an id based on their location.

I also created a separate package of functions used in the indicated scripts, especially for the most extensive Zone Allocation script. It can be found in the location: *./data_tools*

# Setup

The project's requirements are listed in the *requirements.txt* file, which is available in the project's root directory.

To deploy this project locally on your device, use the "python -m venv <name_env>" command in the desired directory.
Then load the project requirements using the command: "pip install -r requirements.txt".

# Usages

  ## Excel Exporter
  - It exports information about the design elements highlighted in the opened Archicad's window to an Excel file. It is also possible to freely select the set of checked properties. It can be used in tandem with an Excel importer script gaining the ability to edit design elements through an Excel sheet.
  - With the help of variables in the configuration part, you can set the set of analyzed properties and the path of the new Excel file with its name.
  - In the second column of the exported file, you need to include guid numbers, which are unique identifiers for each object in project. 
  - The names of the properties you have selected will be in the second row of the file as column headers starting from the second column. They are needed so that it is possible to quickly edit the exported properties and import them into Archicad using Excel importer.
  - This script is inspired by the script created by Graphisoft with the same name.
  ## Excel importer
  - It takes all the data from an Excel file and sets the properties of the listed objects as specified in such a file.
  - In order for the properties to be imported correctly in the Excel file, the guid's of all the objects must be in the first column, and the guid's of all the properties you want to change must be in the first row.
  - In the configuration area of the script in the variable values, you need to specify: the path of this folder where the file to be imported is located and the name of the imported file.
  - This script is inspired by the script created by Graphisoft with the same name.
  ## Zones numbering
  - It numbers all the zones in the project. The assigned numbers always start with the floor number. The script does not read the floor numbers from the project, but segregates the zones into separate groups based on the height at which they are located. 
  - It assigns the next storey number to each successive group. So if, for example, there is no zone in the middle storey, the script will not take this into account when numbering.
  - The value of the STORY_GROUPING_LIMIT variable determines how big the maximum difference in height between zones (between the 'bottom' of one zone and the other) can be in order for them to be assigned to the same floor. Reported in meters.
  - The value of the storyIndex variable determines the value of the first floor number that will be given to the first lowest group of zones. The value of this variable should be equal to the number of the lowest floor on which the zones are located.
 ## Zones allocation
  - The script checks which elements among those checked are within each zone.
  - On input, it needs an IFC file with the geometry of the zones and specially defined properties. 
  - On the output, it writes an Excel file with a breakdown of the allocations made.
 ### Script needs
  - The project must import all necessary properties from the *others\archicad_properties\zone_allocation.xml* file. 
  - When running the script, an Excel file with the same name and location as the output sheet defined in the algorithm must not be open. However, should this happen, the script will notify you.
  - This scripts used functions defined in *others* folder. To configure the project to your needs, you need to follow well-defined steps.

### Configuration steps
- First, you need to export the zone geometry info to an IFC file. 
To do this, select all analyzed zones and use the 'Save as...' command. 
This is what the settings at the bottom of the window should look like:

![obraz](https://user-images.githubusercontent.com/109000485/228182894-9c666f6e-b1c7-4672-ae17-aea7513b7685.png)

- In the main script file in the SCRIPT CONFIGURATION section, provide the desired information, including the path of the exported IFC file and the path of the new Excel file.
- Make sure that the project has imported properties from the file. If not, import them manually.
- Also make sure that the file you want to import the results into is not open. For the script to work the output file must either not exist or be closed.
- Run the script in your IDE.

### Modeling guidelines
 - One of the required properties called 'Name' specifies the information about an element to appear in the sheet in the 'Material' column if you can't pull layer structure, profile or material information from the element. Therefore, if you put an element from which you want to pull information, and its type is in the ![Exclusions](#exclusions) section of this documentation, then this property contains the exact information about what it is.
 - In order for the script to work correctly for standing partitions that do not have uprights, care must be taken to ensure that the reference line is located exactly in the center of such a partition and that the actual thickness of the partition is specified in the settings in the 'Nominal thickness' field.
 - The script allocates to the zone elements that are at most 1 mm from the edge of the zone. The allocation mechanism works best when the edges of the object are at right angles. If elements such as walls are at a significant angle, then the allocation is slightly less accurate.

### Guide to some script configuration variables

#### TypesOfElements
Specifies which types of elements are to be analyzed. Type names should be entered in Polish and should correspond to the names in Polish-language Archicad with one small exception. For a structural partition, give an abbreviated name: 'partition'. Enter subsequent names after a comma in separate quotation marks. The names are not case sensitive.

#### StatesOfCurtainWall
 The script divides partitions into 2 groups: lying and standing. With this variable you specify which types of partitions you want to analyze. The first value in parentheses means 'Do you consider standing partitions?' and the second: 'Do you consider lying partitions?'
 
 ### Exclusions
 Object types that are not taken into account by the script when allocating elements:
 * Stair
 * Railing
 * CurtainWall
 * Door
 * Window
 * Skylight
 
## IDs assignment
  - The script assigns all elements on the exposed layers an Id consisting of two parts: a letter (denoting the type of element) and a number (which element of a given type it is in turn).
  - When the script encounters at least 2 objects with the same dimensions, it assigns them all the Id of the one that was put up the earliest.
  - The walls are divided by the script into 3 groups: external, internal and unidentified. This allocation is carried out based on the value of the "Location" property displayed in the "ID and Category" group. It must be completed for the assignment to be carried out correctly. Depending on the location so expressed, the wall ID starts with the prefix "SZ" (external), "SW" (internal) or "SN" (unidentified).
  - Make sure that all items that should be assigned a new Id are visible. The script will not change the Id of objects which are not showed on the current plan.

## Room report
  - For the script to execute, all properties from the *others\archicad_properties\room_report.xml* file must be available in the project. It must be imported into the project before running the script. 
  - You must also export the model to an IFC file and define all the necessary project parameters before starting. 
  - From the 'Project Info' space, the parameters 'Project - Name', 'Project-Phase', 'Project - Investment Address' 'Designer Full Name' will be imported.  
  - The user should also add 3 IFC properties to the project: 'project_number', 'start_date' and 'end_date'.
  - The script processes the data and produces a report from it in a new DOCX-formatted file, in which it includes general information about the project and information about the rooms, along with a list of the objects in each room.
  - In each list, duplicates are aggregated, i.e. objects occurring at least 2 times are not listed multiple times, but only once, with the sequence 'x <number_of_appearances>' appended to the name.

## Shared Id
  - When the script encounters at least 2 objects with the same dimensions, it assigns them all the Id of the one that was put up the earliest.
  - All analyzed objects must be on visible layers.
  - The script does not analyze shapes and walls.

## Shared Id for walls
  - Assigns a common Id to walls that share a layered structure. 
  - The Id of the earliest placed wall with the given structure is assigned.
  - All analyzed walls must be on visible layers.
  - The script analyzes only the walls.

## Floor Space Distribution
  - Properties from the *others/archicad_properties/placement.xml* file must be available for the script to run in the project. These must be imported before running it.
  - Using the 'precision' variable, the user can set the rounding of the values shown on the charts. 
  - On the other hand, the 'checked_stories' variable defines names of the floors that will be analyzed. A separate chart will be generated for each of these floors. If this variable is empty, then the script will subject all floors to analysis.
  - The 'plot_folder' stores the path to the folder where the charts are to be stored. By default, it is *Other/charts*.
  - Using the 'prefix_mode' variable, the user determines whether he wants to see the name of the zone together with its number or the name alone on the chart. Taking both parameters together is necessary in case the names of the rooms repeat in the BIM model. Otherwise, the Matplotlib library will overwrite their previous 
  occurrences and the generated drawing will be incorrect.
  - For each of the zones located on a given floor, the percentages of the share of the total area accounted for by all zones on that floor are assigned. If the shares of the smallest of the zones are less than 5%, then a bar chart will be generated, otherwise a pie chart will be produced.
  - The files that are the end result of the script are automatically saved with the floor name to PNG format.

## Table of furniture
  - It creates and saves new Excel files in location specified by two user defined variables contains
  file name and its location. 
  - It gets values from properties given in "properties" variable. Properties stated in the brackets are user defined. Make sure they are all available in the project. The names of the properties are given in Polish. You can change it if necessary by modifying the contents of this variable. 
  - The script creates a new sheet, where it includes the results of the assignments made and the property values taken from Archicad. Viewed from the left, the following are placed: the name and number of the zone associated with a given group of objects; the object's id, its name, the dimensions composed of the concatenated values of the individual dimensions, the quantity, the unit price and the summed value of all the units of a given object. 
  - Headers of columns are default in Polish language, but can be modified by setting variable named "titles" 
  - Also calculated is the sum of the price value of all objects in each room and the total value of all those included in the sheet. Summaries go into the sheet not as values, but as summation formulas that can be used to quickly check the price values of different design variants.
 
# Project Status

Project is: done. But the quality of the code is currently being improved.

# Acknowledgements

Many thanks to Barbara Kowalczyk, for the idea of this project and support shown during the implementation.

# Contact

Created by @Marcin-Ramotowski - feel free to contact me!
