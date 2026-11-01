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

## 🎯 MITRE ATT&CK Mapping

* **Execution:** T1204.002 (User Execution: Malicious File)
* **Defense Evasion:** T1140 (Deobfuscate/Decode Files or Information)
* **Command and Scripting Interpreter:** T1059.005 (Visual Basic)
