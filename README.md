Score Calculator 🧮
Welcome to the Score Calculator! This is a simple application designed to help you manage and score any kind of competition, from a science fair to a bake-off. You can create categories, weigh them differently, and automatically calculate final scores for all your submissions.

Features ✨
Easy Data Entry: Add or edit submissions in a simple form.
Custom Scoring Rules: Define your own scoring categories, set custom weights, and even apply a "curve" to the scores.
Advanced Calculations: Apply pre-calculation functions like "Square Root," "Z-Score," or "Rank Order" to scores before they are finalized.
Save & Load: Save your entire competition (settings and all submissions) to a file and load it back up later.
Generate Reports: Instantly generate a final leaderboard and see category winners.
Export to Word: Create a professional, formatted .docx report of the final results with a single click.






Installation and First-Time Setup Guide 🚀
This guide will walk you through every single step to get the application running. No prior coding experience is needed! We'll go slow and explain everything.



Step 1: Install Python on Your Computer
The Score Calculator is written in a language called Python. To run it, you first need to have Python installed on your computer.

How to Check if You Already Have Python:
Open the command line tool:
On Windows: Press the Windows Key (the one with the logo), type cmd, and press Enter. A black window called "Command Prompt" will open.
On Mac: Open your "Applications" folder, then open the "Utilities" folder. Double-click on "Terminal".
Type the following command into the black window and press Enter: "python --version"
If you see text that says something like Python 3.8.10 (any number starting with 3 is great!), you can skip to Step 2.
If you see an error message like "command not found" or the version number starts with a 2, you need to install Python.

How to Install Python:
Go to the official Python download page: https://www.python.org/downloads/
Click the big yellow button to download the latest version for your operating system (Windows or Mac).
Run the installer you just downloaded.
VERY IMPORTANT (for Windows users): On the first screen of the installer, make sure you check the box at the bottom that says "Add Python to PATH". This will make everything much easier later!
Click "Install Now" and follow the on-screen instructions to complete the installation.



Step 2: Download the Application Files
Now you need to extract the files for the Score Calculator itself.
Create a new folder somewhere you can easily find it, like your Desktop. Name it something like ScoreApp.
Extract all 4 files you just downloaded into that folder.



Step 3: Install the Helper Program (OPTIONAL)
The app has one optional helper program it needs to export reports to Microsoft Word. We'll install it using a tool called pip that came with your Python installation.
Open the command line tool (Command Prompt or Terminal) just like you did in Step 1.
Navigate to your application folder. This is the trickiest part. You need to tell the command line where your ScoreApp folder is. You'll use the cd (Change Directory) command.
For example, if your ScoreApp folder is on your Desktop, you would type this command and press Enter: "cd Desktop/ScoreApp"
Tip: You can often drag the folder from your file explorer directly into the command line window to paste its location!
Once you are in the correct folder, run the installation command. Type the following command and press Enter: "pip install -r requirements.txt"
You should see some text as it downloads and installs the python-docx library. This gives the app the power to create Word documents.



Step 4: Run the Score Calculator! 🎉
You've done all the hard work! Now you can run the application.

Make sure you are still in your ScoreApp folder in the command line tool (Command Prompt or Terminal).
Type the following command and press Enter: "python Scoring.py"
The Score Calculator application window should now appear on your screen!

ALTERNATIVE: You can also just run "Competition Scoring.bat" to start the application







How to Use the Application
Here’s a quick tour of how to use the Score Calculator.

The Main Screen
The main screen is split into two parts:

Left Side: A list of all the submissions you've entered.

Right Side: A form to add a new submission or view the details of a selected one.

⚙️ Settings (The Most Important Part!)
Before you start adding submissions, you should visit the settings. Go to File > Settings... in the top menu.

Competition Name: Give your event a title.

Enable Score Curve: This is a powerful feature. When checked, it finds the highest score in each category and scales it up to be the maximum possible score (e.g., 100). All other scores in that category are scaled up by the same percentage. This helps reward people who outperform others.

Enable Custom Weights: Check this if you want some categories to be more important than others (e.g., "Taste" is 70% of the score and "Presentation" is 30%). If you don't check this, all categories will be weighted equally.

Score Scale: The minimum and maximum raw score a judge can give (e.g., 1 to 100).

Categories:

Number of Categories: How many things you will be scoring.

Category Name: The name of each scoring metric (e.g., "Combat," "Design," "Speed").

Weight (%): If custom weights are enabled, enter the percentage for each category. They must all add up to 100!

Pre-Calculation: An advanced option to transform scores. For example, "Invert" is great for races where a lower time is a better score. Hover over the (?) for a full explanation of each option.



Adding and Editing Submissions
To add a new submission, make sure the form on the right is empty (click the ➕ Add New button if it's not).

Fill in the Submission Name, the Scores for each category, and any Notes.

Click the 💾 Save button. The submission will appear in the list on the left.

To edit a submission, simply click on it in the list, change the details in the form on the right, and click 💾 Save again.

📊 Generating the Report
Once all your submissions are entered, click the 📊 Generate Report button at the bottom.

A new window will appear showing the final, ranked leaderboard.

It shows each submission's Rank, Name, Final Score, and a breakdown of their raw score vs. their curved score for each category.

At the bottom of the report, you can click 📄 Export to Word to save a beautifully formatted document of the results.

Saving Your Work
Don't forget to save your whole competition! Use File > Save Competition As... to save a .json file that contains all your settings and submissions. You can open it again later using File > Load Competition....