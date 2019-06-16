#  OpenVPN configuration file CONVERTER for OpenWRT

 This will CONVERT your ".OVPN" configuration files from "<vpn config file>.ovpn" to something, which
 OpenWRT can work with. You can select (within this code) to create single CONF files for each of the OVPN files. (deafult is ONE BIG config file)

 The current configuration will convert the ".ovpn" files into a config similar to "/etc/config/openvpn" and merge this config file with your excisting "/etc/config/openvpn" file.
 
 It will create crypto information from any of your ".ovpn" files (e.g. client.key / client.cert etc.) This information is copied out of the ".ovpn" files and put in separate files under a crypto directory

 Once all that is done, it will "SFTP" (Secure Copy from Windows to Router using SSH) all files to the router. All files will land in "/etc/openvpn".
 
 Then it will merge the created "openvpn.conf" files with your "/etc/config/openvpn" file

 ## WHAT YOU NEED:

 - EXCEL
 - current copy of PuTTY & TOOLS (plink / psftp): http://www.putty.org/ (Tested with Version: 0.70)
 - The OpenWRT router Version 18.06 (This is the one I testet this on. Other versions should probably work as well)

    ### WHY EXCEL? 
    It provides an easy, interactive IDE. 
    
    One could wrapp this into a compiled exe ... just need to create  some forms etc. to allow basic interactions (e.g. configure settings / start conversion etc.).
 
     If I would do it ... I would definately have all settings in an external TXT file ... 
     
     I work with Excel almost everyday ... so just adding another botton to do some router magic is not a big deal :)

## GET STARTED:
    
    Create / Open an EMPTY Excel Workbook
    
    Create a SIMPLE Macro within Excel ... just start recording a macro ...change a cell ... stop recording.
    Make sure you have the macro STORED with the current / new workbook.
    
    Once that is done, Exel will have created a MODULE with a simple macro.
    
    Press ALT-11 to invoke the IDE 
    
    look on your LEFT side, select the module1 wich belongs to your just created workbook.
    
    cut and paste the macro code from here into the module1 in excel.
    
    CONFIRM / ADJUST the settings in this VBA module to YOUR environment
    
    Execute the "Convert_OpenVPN_Config_Files" Macro

### License
----

Apache License, Version 2.0
