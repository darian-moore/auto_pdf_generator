# Auto PDF Generator with Selenium

This program was developed to make the tedious task of generating PDFs autonomous. The user will be given 3 options on how they want to generate their
PDFs; by model, by package, or by assembly (all PDFs will be an assembly because models and packages contain assemblies). Based off what option the user 
chooses they will be prompted to enter some information such as model name, package name, and or assembly name. The user also has the option to select 
multiple models, packages, and or assemblies (which can be from different jobs) to be added in the list assemblies to generate a PDF for. Once all information 
for the models, packages, and or assemblies has been entered by the user, selenium will be used to interact with the web page. Selenium will be used to interact
with the website by logging into the website, finding all assemblies needed, and generating a PDF file for each assembly needed. Each PDF generated will be 
re-located from the deafualt download path, the downloads folder, to a specified directory. While the PDFs are being generated the user will stay updated on how 
many assemblies there are in the list, what assembly they are currently on, what model or package they are on if this option is selected, and updates on what 
the program is doing in the web browser as well as the file system. Once the program has completely finished running it will let the user know how long it took 
to complete generating all PDFs and then terminate itself after 10 seconds.

FUN FACT: program is named "blaineus" because I specifically made this for a guy named Blaine who was complaining about downloading PDFs lol
