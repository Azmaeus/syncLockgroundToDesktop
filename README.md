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
I'll fill this out soon, the gist is that you set-up a Windows Task Schedule to run periodically  