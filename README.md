# Analyzing-a-Malicious-Word-Document
Writing      Static analysis of a suspicious Word document (.docm) using oleid and olevba to identify malicious VBA macros, trace the attack flow, extract IOCs, and map findings to MITRE ATT&amp;CK.
# Analyzing a Malicious Word Document (.docm)

A comprehensive hands-on static analysis lab focusing on dissecting a suspicious Microsoft Word macro-enabled document to identify malicious indicators of compromise (IOCs), extract VBA macros, and map behaviors to the MITRE ATT&CK framework.

---




## 📋 Sample Overview

* **Sample ID:** `[ 2dc9ea3c14cb2af3a4eb4cd4688a3540cae434b4e5394ea09b4b64eb6d4a81cc]`
* **File Name:** `file.docm`
* **File Type:** Microsoft Word Macro-Enabled Document (.docm)
* **SHA256 Hash:** `[2dc9ea3c14cb2a4eb4cd4688a3540cae434b4e5394ea09b4b64eb6d4a81cc]`
* **MD5 Hash:** `[d0fab6619f0456203263ab3bc289cb40]`

---

## 🛠️ Tools Used

* **Kali Linux:** Primary analysis environment.
* **oletools (olevba / oleid):** For detecting VBA macros and analyzing malicious streams within Office documents.
* **CyberChef:** For deobfuscating strings, URLs, and payloads.

---

## 🔍 Analysis Steps & Findings

1. **File Identification & Hashing:**
   * Navigated to the working directory in Kali Linux and calculated cryptographic hashes (`SHA256` and `MD5`) to ensure sample integrity and uniquely identify the artifact.

2. **Static Analysis & Macro Extraction:**
   * Ran `oleid` to check for suspicious indicators (such as VBA macros, embedded objects, or auto-execution triggers).
   * Used `olevba` to extract and inspect Visual Basic for Applications (VBA) source code embedded within the document.

3. **Deobfuscation & IOC Extraction:**
   * Examined obfuscated strings and concatenated variables inside the macro code.
   * Leveraged CyberChef to clean up string manipulations, revealing suspicious command-line execution patterns or dropped payloads.

---

## 🎯 MITRE ATT&CK Mapping

* **Execution:** T1204.002 (User Execution: Malicious File)
* **Defense Evasion:** T1140 (Deobfuscate/Decode Files or Information)
* **Command and Scripting Interpreter:** T1059.005 (Visual Basic)]`
* **File Name:** `file.docm`
* **File Type:** Microsoft Word Macro-Enabled Document (.docm)
* **SHA256 Hash:** `[2dc9ea3c14cb2af3a4eb4cd4688a3540cae434b4e5394ea09b4b64eb6d4a81cc]`
* **MD5 Hash:** `[d0fab6619f0456203263ab3bc289cb40]`

---

## 🛠️ Tools Used

* **Kali Linux:** Primary analysis environment.
* **oletools (olevba / oleid):** For detecting VBA macros and analyzing malicious streams within Office documents.
* **CyberChef:** For deobfuscating strings, URLs, and payloads.

---

## 🔍 Analysis Steps & Findings

1. **File Identification & Hashing:**
   * Navigated to the working directory in Kali Linux and calculated cryptographic hashes (`SHA256` and `MD5`) to ensure sample integrity and uniquely identify the artifact.

2. **Static Analysis & Macro Extraction:**
   * Ran `oleid` to check for suspicious indicators (such as VBA macros, embedded objects, or auto-execution triggers).
   * Used `olevba` to extract and inspect Visual Basic for Applications (VBA) source code embedded within the document.

3. **Deobfuscation & IOC Extraction:**
   * Examined obfuscated strings and concatenated variables inside the macro code.
   * Leveraged CyberChef to clean up string manipulations, revealing suspicious command-line execution patterns or dropped payloads.

---
In this lab, I analyzed a suspicious Word document (.docm) that I downloaded from Abuse.ch.

I wanted to understand what the document was hiding and what could happen if someone opened it.

I did not open the file in Microsoft Word or run its macros. Instead, I used Kali Linux and static analysis tools to examine it safely.

2. Checking the File

I started with oleid to get some basic information about the document.


The results showed:
<img width="975" height="483" alt="image" src="https://github.com/user-attachments/assets/9a656eab-61ef-4b8b-84d9-b37dd55a8db0" />

- The file was a Macro-Enabled Word Document.

- The document was not encrypted.

- VBA macros were present.

- The macro indicator showed HIGH risk.

This was a reason to investigate the file further, but it was not enough on its own to confirm that the document was malicious.

3. Extracting the VBA Code

Next, I used olevba to extract the VBA code from the document.
<img width="975" height="655" alt="image" src="https://github.com/user-attachments/assets/cceab8de-1eef-4735-8f8c-941a722655e8" />

While reviewing the code, I found several suspicious actions.

First: Automatic Execution

The code contained a function called:

AutoOpen

This function can run automatically when a Word document is opened, if macros are allowed.

It then calls another function named ExecutePayload.

Second: Hiding the Download Link

I found that the code tried to get a URL from the document's hidden properties:

BuiltInDocumentProperties("Comments")

This makes the download link less obvious when someone looks at the code.

Third: Downloading a File

The macro uses:

MSXML2.XMLHTTP

to connect to a remote URL and download a file.

It then uses:

Environ("TEMP")

to save the downloaded file in the Windows temporary folder with an .hta extension.

Fourth: Running the Downloaded File

Finally, the code uses:

WScript.Shell
to launch:
mshta.exe
This Windows utility can be used to run HTA files.
The parameters 0, False indicate an attempt to run the process without showing a window or waiting for it to finish.
4. Understanding the Attack
After reviewing the code, I was able to understand the intended attack sequence:
 - The victim receives a suspicious Word document.
 - The victim opens the document and enables macros.
 - The AutoOpen function starts.
 - The macro retrieves a hidden download URL.
 - It connects to the remote server.
 - It downloads an HTA file.
 - It attempts to run the file using mshta.exe.
This showed me how a simple Word document can be used to start a much larger attack.
5. MITRE ATT&CK Mapping
I mapped the behaviors found in the code to the MITRE ATT&CK framework.
Technique ID         Technique          Description

T1204.002         Malicious File        Using a malicious document

T1059.005         Visual Basic           Executing VBA code

T1105         Ingress Tool Transfer     Downloading an external file

T1218.005          Mshta                Using mshta.exe to run an HTA file

These techniques are based on the code analysis, not on a live attack observation.

6. Indicators of Compromise (IOCs)

During the investigation, I identified the following indicators:

Indicator                          Value

File Extension                     .docm
Macro Function                    AutoOpen
Main Function                  ExecutePayload
HTTP Component                 MSXML2.XMLHTTP
Download Folder                   %TEMP%
Downloaded File                      .hta
Execution Tool                    mshta.exe
Script Object                    WScript.Shell

File Hashes:

SHA256: <img width="647" height="67" alt="image" src="https://github.com/user-attachments/assets/94d11a41-6926-4133-bd1e-a6cd95e25a1e" />
]
MD5: <img width="975" height="74" alt="image" src="https://github.com/user-attachments/assets/938e6f12-f8cb-4000-93fc-4f12a5d80a8a" />



7. How Can a SOC Analyst Detect This Activity?

After understanding how the code works, I looked at how a SOC analyst could detect similar behavior on a Windows machine.

Some useful things to monitor include:

1. Process Creation

Look for suspicious activity such as:

- WINWORD.EXE → mshta.exe

Useful Windows events:

- Sysmon Event ID 1 – Process Creation.

- Security Event ID 4688 – Process Creation.

2. Network Connections

Monitor unusual outbound connections from Office applications or related processes.

Useful sources:

- Sysmon Event ID 3.

- Firewall Logs.

- Proxy Logs.

- Microsoft Defender telemetry.

3. File Creation

Monitor the creation of unexpected .hta files in temporary folders.

Useful event:

Sysmon Event ID 11 – File Create.

These are suggested detection opportunities. They were not directly observed during this static analysis lab.

8. Conclusion

This lab helped me understand how a malicious Word document can be used to download and attempt to run another file on a victim's machine.

I started with oleid, extracted the VBA code using olevba, and followed the instructions to understand the attack step by step.

I also learned how to identify useful indicators, connect suspicious behaviors to MITRE ATT&CK, and think about how a SOC analyst could detect similar activity using Windows event logs.

The most important lesson was that analyzing a suspicious file is not just about finding malicious code. It is also about understanding what the code is trying to do and how its activity could be detected.

9. Tools Used

- Kali Linux

- VirtualBox

- oleid

- olevba

- Abuse.ch

- MITRE ATT&CK


Disclaimer

This project is for educational and defensive cybersecurity purposes.

The sample was analyzed without executing its macros or payload. The original sample is not included in this repository.


