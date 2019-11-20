NAME: Meterpreter_Defender

TYPE: VBS Script

PRIMARY LANGUAGE: VBScript
 
AUTHOR: Justin Grimes

ORIGINAL VERSION DATE: 11/18/2019

CURRENT VERSION DATE: 11/20/2019

VERSION: v1.2


DESCRIPTION: An application for detecting and defending against in-memory hacking tools and exploitations. 
Specifically made to detect Meterpreter based payloads.
Reports information about potential Meterpreter payload detection via email and log files.
Destroys detected meterpreter settings upon detection.





PURPOSE: To detect and mitigate Meterpreter based payloads automatically and with visibility.




INSTALLATION INSTRUCTIONS: 
1. Install Meterpreter_Defender into a subdirectory of your Network-wide scripts folder.
2. Open Meterpreter_Defender.vbs with a text editor and configure the variables at the start of the script to match your environment.
3. Open sendmail.ini with a text editor and configure your email server settings.
4. Run the script automatically on domain workstations at machine startup as SYSTEM with a GPO.



NOTES: 
1. This script MUST be run with administrative rights.
2. If this script is started in regular user mode, it will prompt for administrator elevation.
3. "Fake Sendmail for Windows" is required for this application to send notification emails. Per the "Fake Sendmail" license, the required binaries are provided.
4. To reinstall "Fake Sendmail for Windows" please visit  https://www.glob.com.au/sendmail/
5. Use absolute UNC paths for network addresses. DO NOT run this from a network drive letter. The restartAsAdmin() function will not work properly.
6. If using as a startup/logon script it is advised to use a conditional that checks for the prescence of the script prior to running it.
7. "Meterpreter_Payload_Detection.exe" by Damon Mohammad Bagher is required and included with this application. 
8. To reinstall "Meterpreter_Payload_Detection.exe" please visit https://github.com/DamonMohammadbagher/Meterpreter_Payload_Detection
9. For a really interesting blog post about detecting Meterpreter Payloads, please visit https://www.linkedin.com/pulse/detecting-meterpreter-undetectable-payloads-scanning-mohammadbagher/?trk=pulse_spock-articles
