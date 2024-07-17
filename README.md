# AimLoader
Google Sheets extension to automagically update your AimLab benchmark scores

Note: the api key used in both of these is from a dummy account that is not on leaderboards -> will try to update with a new copy 

## Setup
1. Open your benchmark sheet
2. Open the extensions tab and click on apps script
3. Copy the code in Aimlab.gs or Kovaaks.gs to a new script
4. Save the script and refresh the progress sheet

## Usage
1. Open the Aim Loader tab in the top menu
2. Read setup instructions and complete setup
3. Scores should update every minute, click reload to update immediately

- Note: It's probably a good idea to delete the extension once you're done using it and are switching to a different sheet or something. Google only lets you make x amount of api calls a day, so having it run on 5 different sheets might hit that limit. I'm working on figuring out a way to bypass this
