import pandas as pd
import os

# File and Sheet Variables
image_path = os.path.join(os.path.dirname(__file__), "Company Logo.png")
signature = f'''
            <br><span style= 'color: #E476F44; font-size: 22pt'>David Ryan Lynch</b><br></span>
            PH: +1 706-481-2635<br>
            T & E Specialist<br>
            Home Office<br><br>
            <img src="{image_path}" alt="Company Logo">
            '''