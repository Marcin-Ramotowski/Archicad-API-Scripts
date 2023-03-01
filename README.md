# Archicad-API-Scripts

    Scripts to automate some operations in Archicad

Table of Contents

    General Info
    Technologies Used
    Features
    Setup
    Usages
    Project Status
    Room for Improvement
    Acknowledgements
    Contact

General Information

A set of Python scripts that connect to an open building project
in Archicad using the Archicad Automation API,
retrieve information from it,
and make changes to model elements.

Technologies Used

    Python                version 3.10
    archicad              version 24.3000
    openpyxl              version 3.0.7
    et-xmlfile            version 1.1.0


Features

List the ready scripts here:
    
    Scripts printing information on the screen
      Navigator Elements - It displays a tree showing the available elements in the navigator and their hierarchy.
      Project Elements - It subtotals the elements of each type in the project and prints the results on the screen.
      Repeating IDs - It displays id values that occur more than once in the project.
      Sorting - It sorts elements of a given type according to the first specified value in ascending or descending order.
      Check Properties - It displays values of all selected properties for all elements of the specified type.
      All properties - It displays names of all properties available in the project.
      Get property guid - It gets guid of selected property and displays on the screen.
    Scripts printing information on the screen
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

Setup

The project's requirements are listed in the requirements.txt file, which is available in the project's root directory.

To deploy this project locally on your device, use the "python -m venv <name_env>" command in the desired directory.
Then load the project requirements using the command: "pip install -r requirements.txt".

Usages
  Work in progress. Actually, a more detailed description of all the scripts can be found in the file Wytyczne do skryptów.xlsx but for now only in the Polish language version.

Project Status

Project is: done. But the quality of the code is currently being improved.
Room for Improvement

Room for improvement:

    improve code in IDs assignment script - replacing the structure with a large number of if statements on the dictionary
    use of recursion in Navigator Elements script
    change polish name to english in order to code unification

Acknowledgements

    Many thanks to Barbara Kowalczyk, for the idea of this project and support shown during the implementation.

Contact

Created by @Marcin-Ramotowski - feel free to contact me!
