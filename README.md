# ZOS-API_Standalone_Example_Forum_0
This repository contains my answer to a post in the [Zemax Community Forum](https://my.zemax.com/en-US/forum/threads/465e875b-9bce-eb11-bacc-00224809a5b8). The original request was:

---
*I am new to the Python-Zemax ZOS-API so it might be better to do a standalone as you suggest, also in the long run. I have already written some code in the Zemax macro, so I thought I would need to learn fewer new commands using the inherent mode.*

*I also thought that in the specific case of what I want to do, it would be more compact:*

*1. The macro runs the POP for different configurations and saves the data*
*2. It calls Python to perform calculations with the saved data and imports some output variables from Python*
*3. It uploads the values to a merit function*
*4. This macro will be used also for tolerancing with merit function . The merit function will call the script.*

*I would appreciate also recommendations of articles or starting code with the standalone application.*

---
My answer consists in the ZOS-API standalone Python script: **standalone_example_for_jose.py**, which runs on the Double Gauss sample file (installed by default). It requires the macro: **Standalone_text_example.ZPL** to run properly. I didn't quite understand 4. but I'll ask for more details. The **OpticStudio_StandaloneExample.yaml** file should be the Python environment required for this application, but I haven't tested it. The code has a reference to this [GIST](https://gist.github.com/Omnistic/cfff35796e7cf9cbbda9b1de90d104e2) in the comment, which illustrates how to use POPD to retrive text information from a POP analysis window.
