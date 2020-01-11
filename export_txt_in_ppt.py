import win32com
from win32com.client import Dispatch, constants

ppt = win32com.client.Dispatch('PowerPoint.Application')
ppt.Visible = 1
ppt_path = r"C:\Users\zhaoynh\PycharmProjects\python_study\real_project\automatic_report\ppt_text_template.pptx"
pptSel = ppt.Presentations.Open(ppt_path)
f = open('1.txt', "w")
slide_count = pptSel.Slides.Count
for i in range(1, slide_count + 1):
    shape_count = pptSel.Slides(i).Shapes.Count
    print(shape_count)
    for j in range(1, slide_count + 1):
        if pptSel.Slides(i).shapes(j).HasTextFrame:
            s = pptSel.Slides(i).Shapes(j).TextFrame.TextRange.Text
            print(s)
            f.write(s + "\n")
s = pptSel.Slides(1).Shapes(4).TextFrame.TextRange.Text
print(pptSel.Slides(1).shapes(1).HasTextFrame)
print(s)
f.close()
ppt.Quit()
