# IT-Dashboard Challenge

This is a repository for IT-Challenge developed with Python, [rpaframework](https://rpaframework.org/releasenotes.html).
### Features

1. **Automate Scrap the** [IT-Dashboard](https://itdashboard.gov/)
2. It will scrap agencies with amount and save it into the Excel file
3. It will open one of the agency and will scrap the Individual Investments into another excel file.
4. It will check in the Individual Investments table that if the **UII** contains link it will open that link and download the PDF file associated with **Download Business Case PDF** button into output folder  
5. Can be test on [robocorp](https://cloud.robocorp.com/)
6. All downloaded PDF's and Excel sheets will be land in **output** folder
7. It will read the downloaded PDF files and get the **Section A** from each PDF then it will compare the values "Name of this Investment" with the column "Investment Title", and the value "Unique Investment Identifier (UII)" with the column "UII"

## Setup (robocorp)

1. [rpaframework](https://rpaframework.org/releasenotes.html)


## File Structure

### [config.py]
All the configurations for process to run.

### [task.py]
It will start the ITDashboard from [process.py] instance and call the 
required functions to perform the challenge.

### [process.py]
This File contains all the code that perform all the required task.

### [conda.yaml]
Having configuration to set up the environment and [rpaframework](https://rpaframework.org/releasenotes.html) dependencies.

### [robot.yaml]
Having configuration for robocorp to run the [conda.yaml] and execute the task.py


You can find more details and a full explanation of the code on [Robocorp documentation](https://robocorp.com/docs/development-guide/browser/rpa-form-challenge)
