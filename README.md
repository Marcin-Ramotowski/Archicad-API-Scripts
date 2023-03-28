# Archicad-API-Scripts

    Scripts to automate some operations in Archicad

# Table of Contents

    General Info
    Technologies Used
    Features
    Setup
    Usages
    Project Status
    Room for Improvement
    Acknowledgements
    Contact

# General Information

A set of Python scripts that connect to an open building project
in Archicad using the Archicad Automation API,
retrieve information from it,
and make changes to model elements.

# Technologies Used

    Python                version 3.10
    archicad              version 24.3000
    openpyxl              version 3.0.7
    et-xmlfile            version 1.1.0


# Features

List the ready scripts here:
    
    Scripts printing information on the screen:
      Navigator Elements - It displays a tree showing the available elements in the navigator and their hierarchy.
      Project Elements - It subtotals the elements of each type in the project and prints the results on the screen.
      Repeating IDs - It displays id values that occur more than once in the project.
      Sorting - It sorts elements of a given type according to the first specified value in ascending or descending order.
      Check Properties - It displays values of all selected properties for all elements of the specified type.
      All properties - It displays names of all properties available in the project.
      Get property guid - It gets guid of selected property and displays on the screen.
    Scripts that set properties of project elements:
      Excel exporter - It creates a new .xlsx file and places in it the values of the indicated properties from the elements of the indicated types.
      Excel importer - It retrieves data from the specified .xlsx file and sets property values in the project according to that data.
      Chair numbering - It numbers the places by grouping them in rows according to height. Great to numbering places on the huge stadium, cinema or theatre.
      Zones numbering - It numbers the zones by assigning them an id based on their location.
      Zone allocation - It assigns objects to the zones within which they are located based on data from the project and the exported IFC file.
      IDs assignment - It assigns unique ids for objects according to Archicad's default ID allocation rules. Great for cleaning up the mess of item IDs.
      Room report - It collects data on the zone representing the room in question and inserts it into the indicated file at the indicated locations.
      Shared Id - It assigns to objects with the same dimensions the id of the first of these twin elements.
      Shared Id for walls - It assigns the same id for these walls which have the same construction composite.
    I also created a separate package of functions used in the indicated scripts, especially for the most extensive Zone Allocation script. It can be found in the location: "Ustawiające właściwości\Pakiet"

# Setup

The project's requirements are listed in the requirements.txt file, which is available in the project's root directory.

To deploy this project locally on your device, use the "python -m venv <name_env>" command in the desired directory.
Then load the project requirements using the command: "pip install -r requirements.txt".

# Usages

  ## Excel Exporter
  By default, it exports information about the dimensions of walls and beams to an Excel file. However, it is possible to freely select the set of analyzed design elements and the set of checked properties. It can be used in tandem with an Excel importer script gaining the ability to edit design elements through an Excel sheet.
  With the help of variables in the configuration part, you can set the considered types of elements, the set of analyzed properties, the path of the new Excel file and its name, in turn.
  In the first column and the first row of the exported file, you need to include guid numbers, which are unique identifiers for each object and each property. They are needed so that it is possible to quickly edit the exported properties and import them into Archicad using Excel importer.
  This script is based on a script created by Graphisoft with the same name.
 ## Excel importer
 It takes all the data from an Excel file and sets the properties of the listed objects as specified in such a file.  In order for the properties to be imported correctly in the Excel file, the guid's of all the objects must be in the first column, and the guid's of all the properties you want to change must be in the first row.
 In the configuration area of the script in the variable values, you need to specify: the path of this folder where the file to be imported is located and the name of the imported file.
 This script is based on a script created by Graphisoft with the same name.
 ## Chair numbering
 Number all objects classified as chairs by assigning them an Id consisting of: the row designation (rows are designated by letters A-Z), the object number in the row, and the name of the side (left or right)
 For objects to be given a row designation other than A, successive rows must be higher and higher by a value not exceeding the value of the ROW_GROUPING_LIMIT variable (measured in meters).
 This script is based on a script created by Graphisoft with the same name.
 ## Zones numbering
 It numbers all the zones in the project. The assigned numbers always start with the floor number. The script does not read the floor numbers from the project, but segregates the zones into separate groups based on the height at which they are located. It assigns the next storey number to each successive group. So if, for example, there is no zone in the middle storey, the script will not take this into account when numbering.
 The value of the STORY_GROUPING_LIMIT variable determines how big the maximum difference in height between zones (between the 'bottom' of one zone and the other) can be in order for them to be assigned to the same floor. Reported in meters.
 The value of the storyIndex variable determines the value of the first floor number that will be given to the first lowest group of zones. The value of this variable should be equal to the number of the lowest floor on which the zones are located.
 This script was created entirely by me.
 ## Zones allocation
The script checks which elements among those checked are within each zone. On input, it needs an IFC file with the geometry of the zones and specially defined properties. On the output, it writes an Excel file with a breakdown of the allocations made.
 ### Script needs
The project must import all necessary properties from the '/Inne/Właściwości/Przydział do stref.xml' file. When running the script, an Excel file with the same name and location as the output sheet defined in the algorithm must not be open. However, should this happen, the script will notify you.
This scripts used functions defined in Pakiet folder. To configure the project to your needs, you need to follow well-defined steps.

### Configuration steps
First, you need to export the zone geometry info to an IFC file. To do this, select all analyzed zones and use the 'Save as...' command. This is what the settings at the bottom of the window should look like:

![obraz](https://user-images.githubusercontent.com/109000485/228182894-9c666f6e-b1c7-4672-ae17-aea7513b7685.png)

In the main script file in the SCRIPT CONFIGURATION section, provide the desired information, including the path of the exported IFC file and the path of the new Excel file.
Make sure that the project has imported properties from the file '/Inne/Właściwości/Przydział do stref.xml'. If not, import them manually.
Also make sure that the file you want to import the results into is not open. For the script to work the output file must either not exist or be closed.
Run the script in your IDE.

### Modeling guidelines
 One of the required properties called 'Name' specifies the information about an element to appear in the sheet in the 'Material' column if you can't pull layer structure, profile or material information from the element. Therefore, if you put an element from which you want to pull information, and its type is in the 'Exclusions' section of this documentation, then this property contains the exact information about what it is.
 In order for the script to work correctly for standing partitions that do not have uprights, care must be taken to ensure that the reference line is located exactly in the center of such a partition and that the actual thickness of the partition is specified in the settings in the 'Nominal thickness' field.
 The script allocates to the zone elements that are at most 1 mm from the edge of the zone. The allocation mechanism works best when the edges of the object are at right angles. If elements such as walls are at a significant angle, then the allocation is slightly less accurate.

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
The script assigns all elements on the exposed layers an Id consisting of two parts: a letter (denoting the type of element) and a number (which element of a given type it is in turn).
For the script to work, you need the 'Położenie' property, which collects data about the position of the walls in order to distinguish exterior walls from interior ones. If you don't have it, you need to import it from the file '/Inne/Właściwości/Położenie.xml'. It gets this information from the Archicad built-in property with the 'Location' name displayed in the 'ID and Category' group.
Walls should have the 'Położenie' property completed. Walls with undefined position will not get a new Id.
Make sure that all items that should be assigned a new Id are visible. The script will not change the Id of invisible objects.

## Room report
For each zone, it generates a room report containing mainly such information as the zone's equipment, a list of doors and windows in the zone, and a list of neighboring zones based on a pattern from the file: '/Inne/Arkusze/RDS Wzór.xlsx'
The first two variables in the configuration part of the script specify, in turn, the path of the newly created sheet and its name.
All properties from the file '/Inne/Właściwości/Dane strefy.xml' must be available:
From the 'Zone Data' group, properties named: 'Wymagana temperatura', 'Wymagane oświetlenie' collect passively (i.e., the user has to set a value for them) the information that will be placed in the report in the upper right corner. If you want to include them in the report, these properties should be completed.

## Shared Id
When the script encounters at least 2 objects with the same dimensions, it assigns them all the Id of the one that was put up the earliest.
All analyzed objects must be on visible layers. The script does not analyze shapes and walls.

## Shared Id for walls
Assigns a common Id to walls that share a layered structure. The Id of the earliest placed wall with the given structure is assigned.
All analyzed walls must be on visible layers. The script analyzes only the walls.
 
# Project Status

Project is: done. But the quality of the code is currently being improved.
# Room for Improvement
    improve code in IDs assignment script - replacing the structure with a large number of if statements on the dictionary
    use of recursion in Navigator Elements script
    change polish name to english in order to code unification

# Acknowledgements

    Many thanks to Barbara Kowalczyk, for the idea of this project and support shown during the implementation.

# Contact

Created by @Marcin-Ramotowski - feel free to contact me!
