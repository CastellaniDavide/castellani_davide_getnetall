# getnetall
![Author](https://img.shields.io/badge/author-Castellani%20Davide-green?style=flat)
![Version](https://img.shields.io/badge/version-v02.01-blue?style=flat)
![GitHub release (latest by date)](https://img.shields.io/github/v/release/CastellaniDavide/castellani_davide_getnetall?label=lastest%20relase)
![Languages vbs](https://img.shields.io/badge/languages-vbs%20&%20python3-yellowgreen?style=flat)
![Platform supported](https://img.shields.io/badge/platform%20supported-Windows%2010-blue?style=flat)
[![On GitHub](https://img.shields.io/badge/on%20GitHub-True-green?style=flat&logo=github)](https://github.com/CastellaniDavide/castellani_davide_getnetall)

## Tags
 #sri, #wmi, #wmic, #gui, #vbs, #python, #python3, #automatitions

## Description
Get the most important network adapter configuration Infos. To see this you can use the new website [https://castellanidavide.github.io/getnetall/](https://castellanidavide.github.io/getnetall/)

## Resources
To make this program I used this resource(s):

  - [https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-networkadapterconfiguration](https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-networkadapterconfiguration)
  - [https://docs.microsoft.com/en-us/previous-versions//aa394084(v=vs.85)?redirectedfrom=MSDN](https://docs.microsoft.com/en-us/previous-versions//aa394084(v=vs.85)?redirectedfrom=MSDN)
  - [https://docs.microsoft.com/it-it/windows/win32/cimwin32prov/win32-networkclient](https://docs.microsoft.com/it-it/windows/win32/cimwin32prov/win32-networkclient)
  - [https://docs.microsoft.com/it-it/windows/win32/cimwin32prov/win32-networkprotocol](https://docs.microsoft.com/it-it/windows/win32/cimwin32prov/win32-networkprotocol)

## Required
  - Windows (I tried with version 1909 build 18363.815)
  - Python3 (I tried with version 3.7 64bit)
    - wmi package to know how can you install it, you can click [here](https://pypi.org/project/WMI/)

---

### Directories structure
 - bin
     - getinfo.py
     - getinfo.vbs
     - getnetall.vbs
 - doc
	 - README.md
     - getnetall.html (made using pandoc)
 - flussi
     - getinfo_client.csv
     - getinfo_server.csv
     - getnetall.csv
 - log
     - log.log

---

### Execution examples
 - getinfo.py (tested on PowerShell & Command Promt & File Explorer)
	 - ./getinfo.py
     - getinfo.py
     - double click on the program
 - getinfo.vbs (tested on PowerShell & Command Promt)
	 - cscript ./getinfo.vbs
     - cscript getinfo.vbs
     - cscript.exe ./getinfo.vbs
     - cscript.exe getinfo.vbs
 - getnetall.vbs (tested on PowerShell & Command Promt)
	 - cscript ./getnetall.vbs
     - cscript getnetall.vbs
     - cscript.exe ./getnetall.vbs
     - cscript.exe getnetall.vbs

#### Attention:
If you didn't put cscript(.exe) in the *.vbs programs, the programs will open lots of windows.
For the *.py program(s) you can't put cscript(.exe), and it won't open a lot of windows.

---

## Changelog

- [02.01_2020-5-18](#02.01_2020-5-18)
- [01.01_2020-5-10](#01.01_2020-5-10)

---

### 02.01_2020-5-18

 - #### Added
    - now the README is online: [https://castellanidavide.github.io/getnetall/](https://castellanidavide.github.io/getnetall/)
    - getinfo files
      - now you can know more about the Network settings (client & server) in your PC
    - log file
    - all outputs are pushed on the screen and in the csv files
        - |program name |getinfo_client.csv|getinfo_server.csv|getnetall.csv|
          |    :---:    |       :---:      |      :---:       |    :---:    |
          | getinfo.py  |         X        |        X         |             |
          | getinfo.vbs |         X        |        X         |             |
          |getnetall.vbs|                  |                  |      X      |

 - #### Changed
    - now manage the log file

---

### 01.01_2020-5-10

 - first version
 - the result are printed on the PC screen

---

Made by Castellani Davide
If you have any problem please contact me:

- [help@castellanidavide.it](mailto:help@castellanidavide.it)
- [GitHub Issue](https://github.com/CastellaniDavide/getnetall/issues)