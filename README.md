AI-Poweered Automation For Presentation
-----------------------------------------

This project automates PowerPoint creation using AI-driven prompts with LLaMA2, ensuring text alignment and other parameters in the templates. 
It supports PDF to PPT conversion and provides a Flask web interface for easy user interaction.

------------------------------------------------------------------------------------------------------------------------------------------------------
File Descriptions:
-------------------

app.py - The main Flask application that provides the web interface.
createppt.py - Script that generates PPT files based on user prompts and aligns the text in predefined templates.
pdftoppt.py - Converts PDF files to PowerPoint presentations.

--------------------------------------------------------------------------------------------------------------------------------------------------------
Running the Project
--------------------

To run the web interface, execute the app.py file
This will start a local web server where users can input prompts to generate PPTs automatically.
To convert PDF files to PPT format, use the pdftoppt.py script which create the ppt based on the content in the file.
To create PPTs directly from prompts, run the createppt.py script

----------------------------------------------------------------------------------------------------------------------------------------------------------
Requirement
--------------

Before running the project, ensure you have the following installed:
Python 3.
Flask
LLaMA2 model via Ollama (runs locally)

----------------------------------------------------------------------------------------------------------------------------------------------------------
As the app.py is the main file which contains the main source file. The first route will lead us to the landing page and then to the login, alse you can move directly to he choice page
where we want to enter our prompt.
The choice page has two different work to do first is "prompt" where the user can create ppt by using the prompt. Here we are using llama model which we run locally in our laptop.
The next one is file to ppt converter where you can uplode some files which turns into a pptx formet.
The createppt.py fiels contains the code which generate the ppt.
pdftoppt.py file contain the code where we convert the files into a ppt formet.
