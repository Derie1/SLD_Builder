# SLD_builder

### Description

SLD_builder is a script for building single line diagram of electrical distribution board from excel colculation file. SLD builds in Autodesk AutoCAD application.

### Video demonstration

[![Watch the video](https://i9.ytimg.com/vi/oWrLLjZIoMI/mq2.jpg?sqp=CMDG65oG&rs=AOn4CLAaAE0nqTkAg_HEYNn_SF7A_RSbKA)](https://youtu.be/oWrLLjZIoMI)

### Usage

First, you should create colculation with excel file (use template from "test_files" folder).
Then, run SLD_builder.exe, select excel file you created on first step, and then select SLD_template.dwg from "test_files" folder to create diagram. After script work done, save your file as you wish.
Later, if you need to take changes in your SLD created before, you need to make changes in excel file. Then run SLD_builder again, select excel file first, and then select your dwg file created before. Script will delete existing lines and circuit breakers and places new ones with chenged data.
