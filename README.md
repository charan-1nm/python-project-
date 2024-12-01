
import math
from openpyxl.chart import LineChart
from openpyxl import Workbook
import os
import tkinter as tk
from tkinter import filedialog
import shutil

# Define your existing code as a function
def process_dot_file(file_path):
# Your existing code here
x = os.path.dirname(file_path)
y = os.path.basename(file_path)
a = &quot;out.text&quot;
shutil.copy(file_path, a)
a1 = &quot;output1.csv&quot;

with open(a, &#39;r&#39;) as b1:
lines = [line.strip().split() for line in b1]
csv_lines = [&#39;,&#39;.join(line) for line in lines]
with open(a1, &#39;w&#39;) as h1:
h1.write(&#39;\n&#39;.join(csv_lines))
os.remove(a)

df = pd.read_csv(a1)


os.remove(a1)
z = df.columns.tolist()
i = 0
s = 0
s1 = 0
p = 0
L = []
L1 = []
L3 = []
L4 = []
M = []
M1 = []
j = math.floor(df.loc[0, &quot;ElActual&quot;]) + 5
H = 5
def G(M1):
l = 0
a = len(M1)
while(True):
k = (l + a) // 2
if M1[k] &lt; M1[k+1]:
l = k + 1
else:
a = k
if a == l:
break


return k
if(&quot;Avg_El_iM1&quot; not in z):
k = z.index(&quot;El_iM1&quot;)
df.insert(k + 1, &quot;Avg_El_iM1&quot;, None)
df.insert(k + 3, &quot;Avg_El_iM2&quot;, None)
kum = G(df[&quot;ElActual&quot;])
while(i &lt;= kum):
if(df.loc[i, &quot;ElActual&quot;] &lt; j):
s = s + df.loc[i, &quot;El_iM1&quot;]
s1 = s1 + df.loc[i, &quot;El_iM2&quot;]
else:
if i - p != 0:
df.loc[i - 1, &quot;Avg_El_iM1&quot;] = s / (i - p)
L.append(format(s / (i - p), &quot;.3f&quot;))
s = 0
df.loc[i - 1, &quot;Avg_El_iM2&quot;] = s1 / (i - p)
L3.append(format(s1 / (i - p), &quot;.3f&quot;))
s1 = 0
p = i
M.append(H)
H = H + 5
j = j + 5
i = i + 1
L.append(format(s / (i - p), &quot;.3f&quot;))
L3.append(format(s1 / (i - p), &quot;.3f&quot;))


M.append(H)
i = kum + 1
p = i
j = math.floor(df.loc[i, &quot;ElActual&quot;]) - 5
s = 0
s1 = 0
while(i &lt; len(df[&quot;ElActual&quot;])):
if(df.loc[i, &quot;ElActual&quot;] &gt; j):
s = s + df.loc[i, &quot;El_iM1&quot;]
s1 = s + df.loc[i, &quot;El_iM2&quot;]
else:
if i - 1 - p != 0:
df.loc[i - 1, &quot;Avg_El_iM1&quot;] = s / (i - p)
L1.append(format(s / (i - p), &quot;.3f&quot;))
s = 0
df.loc[i - 1, &quot;Avg_El_iM2&quot;] = s1 / (i - p)
L4.append(format(s1 / (i - p), &quot;.3f&quot;))
s1 = 0
p = i
M1.append(H)
H -= 5
j = j - 5
i = i + 1
L1.append(format(s / (i - p), &quot;.3f&quot;))
L4.append(format(s1 / (i - p), &quot;.3f&quot;))


M1.append(H)
p = [None] * len(L)
d = pd.DataFrame()
d[&quot;El&quot;] = M + M1
d[&quot;UPM1&quot;] = L + p
d[&quot;DNM1&quot;] = p + L1
d[&quot;&quot;] = [None] * (len(M) * 2)
d[&quot;El2&quot;] = M + M1
d[&quot;UPM2&quot;] = L3 + p
d[&quot;DNM2&quot;] = p + L4

# Specify the directory where you want to save the output file
output_directory = os.path.dirname(file_path)

# Combine the output directory and the filename
b1 = os.path.join(output_directory, &quot;OutPut&quot; + y[:len(y) - 3] + &quot;xlsx&quot;)
d.to_excel(b1, index=False)

# Read data from the Excel file
df = pd.read_excel(b1)

# Create a new Excel writer object
writer = pd.ExcelWriter(b1, engine=&#39;xlsxwriter&#39;)

# Write the DataFrame to the Excel file


df.to_excel(writer, sheet_name=&#39;Sheet1&#39;, index=False)

# Get the xlsxwriter workbook and worksheet objects
workbook = writer.book
worksheet = writer.sheets[&#39;Sheet1&#39;]

# Create a chart object
chart = workbook.add_chart({&#39;type&#39;: &#39;line&#39;})
chart1 = workbook.add_chart({&#39;type&#39;: &#39;line&#39;})

# Define the data ranges for all five series
categories = &#39;=Sheet1!$A$2:$A$&#39; + str(len(df) + 1)
values_upm1 = &#39;=Sheet1!$B$2:$B$&#39; + str(len(df) + 1)
values_dnm1 = &#39;=Sheet1!$C$2:$C$&#39; + str(len(df) + 1)
values_dnm2 = &#39;=Sheet1!$E$2:$E$&#39; + str(len(df) + 1)
values_upm3 = &#39;=Sheet1!$F$2:$F$&#39; + str(len(df) + 1)
values_dnm3 = &#39;=Sheet1!$G$2:$G$&#39; + str(len(df) + 1)

# Configure the series data
chart.add_series({
&#39;name&#39;: &#39;UPM1&#39;,
&#39;categories&#39;: categories,
&#39;values&#39;: values_upm1,
&#39;line&#39;: {&#39;color&#39;: &#39;red&#39;},
})


chart.add_series({
&#39;name&#39;: &#39;DNM1&#39;,
&#39;categories&#39;: categories,
&#39;values&#39;: values_dnm1,
&#39;line&#39;: {&#39;color&#39;: &#39;blue&#39;},
})

chart.add_series({
&#39;name&#39;: &#39;EL&#39;,
&#39;y2_axis&#39;: True,
&#39;categories&#39;: categories,
&#39;values&#39;: values_dnm2,
&#39;line&#39;: {&#39;color&#39;: &#39;orange&#39;},
})

chart1.add_series({
&#39;name&#39;: &#39;EL&#39;,
&#39;y2_axis&#39;: True,
&#39;categories&#39;: categories,
&#39;values&#39;: values_dnm2,
&#39;line&#39;: {&#39;color&#39;: &#39;orange&#39;},
})

chart1.add_series({


&#39;name&#39;: &#39;M2UP&#39;,
&#39;categories&#39;: categories,
&#39;values&#39;: values_upm3,
&#39;line&#39;: {&#39;color&#39;: &#39;purple&#39;},
})

chart1.add_series({
&#39;name&#39;: &#39;M2DN&#39;,
&#39;categories&#39;: categories,
&#39;values&#39;: values_dnm3,
&#39;line&#39;: {&#39;color&#39;: &#39;yellow&#39;},
})

# Set chart title and axis labels
chart.set_title({&#39;name&#39;: &#39;2D Line Chart&#39;})
chart.set_x_axis({&#39;name&#39;: &#39;EL&#39;})
chart.set_y_axis({&#39;name&#39;: &#39;Values&#39;})
chart.set_y2_axis({&#39;name&#39;: &#39;Degrees&#39;})

chart1.set_title({&#39;name&#39;: &#39;2D Line Chart&#39;})
chart1.set_x_axis({&#39;name&#39;: &#39;EL&#39;})
chart1.set_y_axis({&#39;name&#39;: &#39;Values&#39;})
chart1.set_y2_axis({&#39;name&#39;: &#39;Degrees&#39;})

# Insert the chart into the worksheet


worksheet.insert_chart(&#39;J5&#39;, chart)
worksheet.insert_chart(&#39;J22&#39;, chart1)

# Close the Excel writer (this saves the file)
writer.close()

# Create a function to select any file
def select_file():
file_path = filedialog.askopenfilename(
title=&quot;Select File&quot;,
filetypes=[(&quot;All Files&quot;, &quot;*.*&quot;)], # Changed filetypes to allow all files
initialdir=os.path.expanduser(&quot;~&quot;)
)

if file_path:
process_dot_file(file_path)
else:
print(&quot;No file selected.&quot;)

# Create a tkinter root window
root = tk.Tk()
root.title(&quot;Select File&quot;)

# Function to display M1 values and graph
def display_m1():


# Add code here to display M1 values and graph in the GUI
pass

# Function to display M2 values and graph
def display_m2():
# Add code here to display M2 values and graph in the GUI
pass

# Create buttons for M1 and M2
m1_button = tk.Button(root, text=&quot;M1&quot;, command=display_m1)
m2_button = tk.Button(root, text=&quot;M2&quot;, command=display_m2)

# Create a button to select the DOT file
select_button = tk.Button(root, text=&quot;Select DOT File&quot;, command=select_file)

# Pack buttons
select_button.pack(pady=10)
m1_button.pack(pady=10)
m2_button.pack(pady=10)

root.mainloop()

Output:


Here we have select and give DOT file as the input file.

Here we are selecting input DOT file.


Here we can observe output file is created.

The Output that we got in excel.
import pandas as pd
import math
from openpyxl.chart import LineChart
from openpyxl import Workbook


import os
import tkinter as tk
from tkinter import filedialog
import shutil

# Define your existing code as a function
def process_dot_file(file_path):
# Your existing code here
x = os.path.dirname(file_path)
y = os.path.basename(file_path)
a = &quot;out.text&quot;
shutil.copy(file_path, a)
a1 = &quot;output1.csv&quot;

with open(a, &#39;r&#39;) as b1:
lines = [line.strip().split() for line in b1]
csv_lines = [&#39;,&#39;.join(line) for line in lines]
with open(a1, &#39;w&#39;) as h1:
h1.write(&#39;\n&#39;.join(csv_lines))
os.remove(a)

df = pd.read_csv(a1)
os.remove(a1)
z = df.columns.tolist()
i = 0


s = 0
s1 = 0
p = 0
L = []
L1 = []
L3 = []
L4 = []
M = []
M1 = []
j = math.floor(df.loc[0, &quot;ElActual&quot;]) + 5
H = 5
def G(M1):
l = 0
a = len(M1)
while(True):
k = (l + a) // 2
if M1[k] &lt; M1[k+1]:
l = k + 1
else:
a = k
if a == l:
break
return k
if(&quot;Avg_El_iM1&quot; not in z):
k = z.index(&quot;El_iM1&quot;)


df.insert(k + 1, &quot;Avg_El_iM1&quot;, None)
df.insert(k + 3, &quot;Avg_El_iM2&quot;, None)
kum = G(df[&quot;ElActual&quot;])
while(i &lt;= kum):
if(df.loc[i, &quot;ElActual&quot;] &lt; j):
s = s + df.loc[i, &quot;El_iM1&quot;]
s1 = s1 + df.loc[i, &quot;El_iM2&quot;]
else:
if i - p != 0:
df.loc[i - 1, &quot;Avg_El_iM1&quot;] = s / (i - p)
L.append(format(s / (i - p), &quot;.3f&quot;))
s = 0
df.loc[i - 1, &quot;Avg_El_iM2&quot;] = s1 / (i - p)
L3.append(format(s1 / (i - p), &quot;.3f&quot;))
s1 = 0
p = i
M.append(H)
H = H + 5
j = j + 5
i = i + 1
L.append(format(s / (i - p), &quot;.3f&quot;))
L3.append(format(s1 / (i - p), &quot;.3f&quot;))
M.append(H)
i = kum + 1
p = i


j = math.floor(df.loc[i, &quot;ElActual&quot;]) - 5
s = 0
s1 = 0
while(i &lt; len(df[&quot;ElActual&quot;])):
if(df.loc[i, &quot;ElActual&quot;] &gt; j):
s = s + df.loc[i, &quot;El_iM1&quot;]
s1 = s + df.loc[i, &quot;El_iM2&quot;]
else:
if i - 1 - p != 0:
df.loc[i - 1, &quot;Avg_El_iM1&quot;] = s / (i - p)
L1.append(format(s / (i - p), &quot;.3f&quot;))
s = 0
df.loc[i - 1, &quot;Avg_El_iM2&quot;] = s1 / (i - p)
L4.append(format(s1 / (i - p), &quot;.3f&quot;))
s1 = 0
p = i
M1.append(H)
H -= 5
j = j - 5
i = i + 1
L1.append(format(s / (i - p), &quot;.3f&quot;))
L4.append(format(s1 / (i - p), &quot;.3f&quot;))
M1.append(H)
p = [None] * len(L)
d = pd.DataFrame()


d[&quot;El&quot;] = M + M1
d[&quot;UPM1&quot;] = L + p
d[&quot;DNM1&quot;] = p + L1
d[&quot;&quot;] = [None] * (len(M) * 2)
d[&quot;El2&quot;] = M + M1
d[&quot;UPM2&quot;] = L3 + p
d[&quot;DNM2&quot;] = p + L4

# Specify the directory where you want to save the output file
output_directory = os.path.dirname(file_path)

# Combine the output directory and the filename
b1 = os.path.join(output_directory, &quot;OutPut&quot; + y[:len(y) - 3] + &quot;xlsx&quot;)
d.to_excel(b1, index=False)

# Read data from the Excel file
df = pd.read_excel(b1)

# Create a new Excel writer object
writer = pd.ExcelWriter(b1, engine=&#39;xlsxwriter&#39;)

# Write the DataFrame to the Excel file
df.to_excel(writer, sheet_name=&#39;Sheet1&#39;, index=False)

# Get the xlsxwriter workbook and worksheet objects


workbook = writer.book
worksheet = writer.sheets[&#39;Sheet1&#39;]

# Create a chart object
chart = workbook.add_chart({&#39;type&#39;: &#39;line&#39;})
chart1 = workbook.add_chart({&#39;type&#39;: &#39;line&#39;})

# Define the data ranges for all five series
categories = &#39;=Sheet1!$A$2:$A$&#39; + str(len(df) + 1)
values_upm1 = &#39;=Sheet1!$B$2:$B$&#39; + str(len(df) + 1)
values_dnm1 = &#39;=Sheet1!$C$2:$C$&#39; + str(len(df) + 1)
values_dnm2 = &#39;=Sheet1!$E$2:$E$&#39; + str(len(df) + 1)
values_upm3 = &#39;=Sheet1!$F$2:$F$&#39; + str(len(df) + 1)
values_dnm3 = &#39;=Sheet1!$G$2:$G$&#39; + str(len(df) + 1)

# Configure the series data
chart.add_series({
&#39;name&#39;: &#39;UPM1&#39;,
&#39;categories&#39;: categories,
&#39;values&#39;: values_upm1,
&#39;line&#39;: {&#39;color&#39;: &#39;red&#39;},
})

chart.add_series({
&#39;name&#39;: &#39;DNM1&#39;,


&#39;categories&#39;: categories,
&#39;values&#39;: values_dnm1,
&#39;line&#39;: {&#39;color&#39;: &#39;blue&#39;},
})

chart.add_series({
&#39;name&#39;: &#39;EL&#39;,
&#39;y2_axis&#39;: True,
&#39;categories&#39;: categories,
&#39;values&#39;: values_dnm2,
&#39;line&#39;: {&#39;color&#39;: &#39;orange&#39;},
})

chart1.add_series({
&#39;name&#39;: &#39;EL&#39;,
&#39;y2_axis&#39;: True,
&#39;categories&#39;: categories,
&#39;values&#39;: values_dnm2,
&#39;line&#39;: {&#39;color&#39;: &#39;orange&#39;},
})

chart1.add_series({
&#39;name&#39;: &#39;M2UP&#39;,
&#39;categories&#39;: categories,
&#39;values&#39;: values_upm3,


&#39;line&#39;: {&#39;color&#39;: &#39;purple&#39;},
})

chart1.add_series({
&#39;name&#39;: &#39;M2DN&#39;,
&#39;categories&#39;: categories,
&#39;values&#39;: values_dnm3,
&#39;line&#39;: {&#39;color&#39;: &#39;yellow&#39;},
})

# Set chart title and axis labels
chart.set_title({&#39;name&#39;: &#39;2D Line Chart&#39;})
chart.set_x_axis({&#39;name&#39;: &#39;EL&#39;})
chart.set_y_axis({&#39;name&#39;: &#39;Values&#39;})
chart.set_y2_axis({&#39;name&#39;: &#39;Degrees&#39;})

chart1.set_title({&#39;name&#39;: &#39;2D Line Chart&#39;})
chart1.set_x_axis({&#39;name&#39;: &#39;EL&#39;})
chart1.set_y_axis({&#39;name&#39;: &#39;Values&#39;})
chart1.set_y2_axis({&#39;name&#39;: &#39;Degrees&#39;})

# Insert the chart into the worksheet
worksheet.insert_chart(&#39;J5&#39;, chart)
worksheet.insert_chart(&#39;J22&#39;, chart1)


# Close the Excel writer (this saves the file)
writer.close()

# Create a function to select any file
def select_file():
file_path = filedialog.askopenfilename(
title=&quot;Select File&quot;,
filetypes=[(&quot;All Files&quot;, &quot;*.*&quot;)], # Changed filetypes to allow all files
initialdir=os.path.expanduser(&quot;~&quot;)
)

if file_path:
process_dot_file(file_path)
else:
print(&quot;No file selected.&quot;)

# Create a tkinter root window
root = tk.Tk()
root.title(&quot;Select File&quot;)

# Function to display M1 values and graph
def display_m1():
# Add code here to display M1 values and graph in the GUI
pass


# Function to display M2 values and graph
def display_m2():
# Add code here to display M2 values and graph in the GUI
pass

# Create buttons for M1 and M2
m1_button = tk.Button(root, text=&quot;M1&quot;, command=display_m1)
m2_button = tk.Button(root, text=&quot;M2&quot;, command=display_m2)

# Create a button to select the DOT file
select_button = tk.Button(root, text=&quot;Select DOT File&quot;, command=select_file)

# Pack buttons
select_button.pack(pady=10)
m1_button.pack(pady=10)
m2_button.pack(pady=10)

root.mainloop()
