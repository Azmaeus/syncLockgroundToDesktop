# syncLockgroundToDesktop
Love those beautiful pictures on the Windows 10 login screen? I do too! This will grab a copy of those pictures and put them in a folder to use for your desktop background.


## How to use

### Configure
Using your favorite text editor, open syncLockgroundToDesktop.vbs  
find the "Configuration" section, has a bunch of "Const" declarations, edit these to your liking  
MIN_PIC_SIZE is the smallest size (in bytes) you want to accept (to ensure the file is likely to be a picture)  
MIN_PIC_WIDTH is the smallest horizontal pixels desired - e.g. 1920 if you only want 1920x1080 or bigger (default is 1680)  
MIN_PIC_HEIGHT is the smallest vertical pixels desired - e.g. 1080 if you only want 1920x1080 or higher (default is 1050)  
LOG_TO_FILE is either True or False, whether you want the script to output what it's doing to file (useful for troubleshooting or verifying it's working, not normally needed)  
LOG_FILE if LOG_TO_FILE is set to True, this is the name (if no path is included in the file name, then the file will be placed in the working directory)  
MAX_LOG_SIZE sets the size (in bytes) limit for each log file  
LOG_BACKUP is either True or False, if False then only the one log file is created, and once it is full it will be replaced entirely  
LOG_FILE_OLD is the name (and path, see LOG_FILE above) of the "backup" log file, when the primary log file hits the MAX_LOG_SIZE  
LOG_TO_SCREEN for a real-time view of what's going on! (if you're into that rather than a log file)  

### Manual use
configure as detailed above  
double-click syncLockgroundToDesktop.vbs  
This should do a one-time copying of anything that's new into the configured folder  


### Automatic use (hands-free, software of the future!)
1. Open Task Scheduler
2. Select the Action menu, and select "Create Basic Task..."
3. Enter the Name and Description as desired (below are what I used), then click Next
    Name: Sync Lockground to Desktop
    Description: Captures "Lockground" (background pictures on the lock screen) images to be used as desktop background wallpapers
4. Select how frequently you want the synchronization to happen then click Next. I prefer to run it nightly when I'm asleep, or first thing when it turns on the next morning, but you can choose what suits you best. This shouldn't take very long, so it shouldn't slow down startup if you choose to run on login or startup. The rest of the guide will presume Daily was chosen, if another option was selected please adjust your steps accordingly.
5. Make the adjustments to the Trigger details, for my "Daily" guide set the Start time to a time you're not normally using the computer. I set mine to 3:00:00 AM, ticked Synchronize across time zones, and Recur every 1 days, then click Next
6. Verify the option "Start a program" is selected, then click Next
7. Click "Browse..." and select the syncLockgroundToDesktop.vbs file
    7a. If logging is enabled, you can either set the LOG_FILE to include a path where the log should be written to, otherwise if you just put a filename you should also in the Task Scheduler include the "Start in (optional)" to point to where the log file should be written to.
8. Click Next, verify the scheduled task is set up as you expect it to be, then click Finish
    8a. Now would be a good time to both test the job just created, and also get the first batch of wallpapers. To do so, select Task Scheduler Library in the left pane. Scroll down in that window until you find the task you just created (named as you did in step 3 above). Right-click that row and select "Run". The Status will change to "Running". You can occasionally press F5 key to refresh, after the task is complete the Status will revert to Ready. There should now be at least one image saved in the location you configured to store the images.
9. Now you're syncrhonizing the background images from the lock screen to a folder you can easily access (and that Microsoft won't periodically delete), you just need to point your wallpapers to that directory.
    9a. Open Windows Settings - Press the Start button, then select the gear-looking icon (when you hover over it it should expand and say "Settings" next to it)
    9b. Click Personalization
    9c. Change the "Background" drop-down menu to be "Slideshow"
    9d. Under "Choose albums for your slideshow" click the "Browse" button
    9e. Navigate to where you have the images saved to, select it, and click "Choose this folder"

There are other settings you can adjust, for example how frequently to change the background, whether to shuffle the order, etc. Set them to your preference, or if you don't know your preference yet leave them alone and come back later when you determine your preference.

That should do it!