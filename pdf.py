from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate,Table,TableStyle, Image
from reportlab.platypus import Paragraph
import textwrap, xlrd, datetime
import plotly.graph_objects as go
from plotly.subplots import make_subplots


#===================== START: Code to read Excel Sheet ==============================


wb = xlrd.open_workbook("./raw_data.xlsx") 
date_convert = wb.datemode
sheet = wb.sheet_by_index(0)
col_heads = sheet.row_values(0)

#===================== END: Code to read Excel Sheet ================================
questions_data = {}

for row in range(1,sheet.nrows):
    stud_id = int(sheet.cell_value(row, 0))
    row_dict = {}
    for col in range(11, sheet.ncols):
        col_obj = sheet.cell_value(row, col)
        row_dict[col_heads[col]] = col_obj
    if stud_id in questions_data:
        questions_data[stud_id].append(row_dict)
    else:
        questions_data[stud_id]= [row_dict]


questions_col = [] #Array of Student Performance Table Heading

for value in col_heads[11:sheet.ncols]:
    wrapper = textwrap.TextWrapper(width=10)
    word_list = wrapper.wrap(text=value)
    temp = ""
    for short in word_list:
        temp += short +"\n"  #Breaking the text if it has width more that 10
    questions_col.append(temp)



marks_data = [] #array That contains all questions and marks
for row in range(1,sheet.nrows):
    each_row = []
    for col in range(sheet.ncols):
        col_obj = sheet.cell_value(row, col)
        each_row.append(col_obj)
    marks_data.append(each_row)


# print(marks_data)
unique_ids = set()

for arr in marks_data:
    if arr[0] not in unique_ids:
        unique_ids.add(arr[0])
        blueprint = Canvas("./pdfs/{}_{}_{}.pdf".format(arr[0],arr[1],arr[2]))
        blueprint.setTitle("Mark Sheet")
        blueprint.setFont("Helvetica-Bold", 16)
        #============================= Begin: Student Details =========================================
        blueprint.drawCentredString(4*inch, 11*inch,"Wisdom Tests and Math Challenge")
        # print(arr[7])
        image_path = "./Pics/{}.jpg".format(int(arr[0]))
        I = Image(image_path)
        I.drawHeight = 1.25*inch*I.drawHeight / I.drawWidth
        I.drawWidth = 1.25*inch
        I.drawOn(blueprint, 6.5*inch, 10*inch)
        data = [
            [Paragraph("<b>Name Of Candidate</b>"),arr[1]],
            [Paragraph("<b>Grade </b>"),arr[3]],
            [Paragraph("<b>School Name</b>"),arr[5]],
            [Paragraph("<b>City Of Residence</b>"),arr[7]],
            [Paragraph("<b>Country Of Residence</b>"),arr[9]],
        ]
        t = Table(data)
        t.wrapOn(blueprint, 4*inch, 10*inch)
        t.drawOn(blueprint, 1*inch, 8.7*inch)
        temp_date = xlrd.xldate_as_tuple(arr[6], date_convert)
        birth_date = (datetime.date(temp_date[0],temp_date[1],temp_date[2])).strftime('%Y-%m-%d')
        temp_date = xlrd.xldate_as_tuple(arr[8], date_convert)
        exam_date = (datetime.date(temp_date[0],temp_date[1],temp_date[2])).strftime('%Y-%m-%d')
        data = [
            [Paragraph("<b>Registration No</b>"),int(arr[2])],
            [Paragraph("<b>Gender</b>"),arr[4]],
            [Paragraph("<b>Date of Birth </b>"),birth_date],
            [Paragraph("<b>Date of Test </b>"),exam_date],
            [Paragraph("<b>Extra Time assistance</b>"),arr[10]],
        ]
        t = Table(data)
        t.wrapOn(blueprint, 4.5*inch, 10*inch)
        t.drawOn(blueprint, 4.5*inch, 8.7*inch)
        #============================= End: Student Details =========================================

        #================================ Begin: Drawing Questions table =====================================
        blueprint.drawCentredString(4*inch, 8.4*inch,"Student Performence")
        if arr[0] in questions_data:
            data = [questions_col]
            question_number = []
            time_taken = []
            attempts_count = [0,0]
            correct_count, wrong_count, total_questions = 0, 0, 0
            for questions in questions_data[arr[0]]:
                question_number.append(questions["Question No."])
                time_taken.append(questions["Time Spent on question (sec)"])
                if questions["Attempt status "] == "Attempted":
                    attempts_count[0] += 1
                else:
                    attempts_count[1] = 1
                if questions["Outcome (Correct/Incorrect/Not Attempted)"] == "Correct":
                    correct_count += 1
                elif questions["Outcome (Correct/Incorrect/Not Attempted)"] == "Incorrect":
                    wrong_count += 1
                total_questions += 1
                temp = []
                for key in questions:
                    if questions[key] is None:
                        temp.append("        ")
                    else:
                        temp.append(str(questions[key]) +"      ")
                data.append(temp)
        t = Table(data,style=[
            ('BACKGROUND', (0, 0), (-1,0), colors.lavender),
            ('BACKGROUND', (0, 1), (-1,-1), colors.pink),
            ('BOX',(0,0),(-1,-1),2,colors.black),
            ('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
        ])
        t.wrapOn(blueprint, 1*inch, 8*inch)
        t.drawOn(blueprint, 0.5*inch, 6.1*inch)
        #================================ End: Drawing Questions table =====================================
        #================================ Begin: Drawing Time Spent PIE Chart =========================================
        fig = go.Figure(data=[go.Pie(labels=question_number, values=time_taken, hole=.3)])
        fig.update_layout( title="Time Spent on Each Question",font=dict(
                family="Courier New, monospace",
                size=18,
                color="RebeccaPurple"
            ),legend=dict(
                x=0,
                y=1,
                traceorder="reversed",
                title_font_family="Times New Roman",
                font=dict(
                    family="Courier",
                    size=16,
                    color="black"
                )
            ))
        fig.write_image("./graphs/{}_time_spent.png".format(arr[0]))
        image_path = "./graphs/{}_time_spent.png".format(arr[0])
        I = Image(image_path)
        I.drawHeight = 4.25*inch*I.drawHeight / I.drawWidth
        I.drawWidth = 4.25*inch
        I.drawOn(blueprint, 0.5*inch, 3*inch)
        #================================ End: Drawing Time Spent PIE Chart =========================================

        #================================ Begin: Drawing Time Spent BAR Graph =========================================
        bar_chart = go.Figure([go.Bar(x=question_number, y=time_taken)])
        bar_chart.update_layout( title="Time Spent on Each Question",font=dict(
                family="Courier New, monospace",
                size=18,
                color="RebeccaPurple"
            ))
        bar_chart.write_image("./graphs/{}_bar_chart_time_spent.png".format(arr[0]))
        image_path = "./graphs/{}_bar_chart_time_spent.png".format(arr[0])
        I = Image(image_path)
        I.drawHeight = 4.25*inch*I.drawHeight / I.drawWidth
        I.drawWidth = 4.25*inch
        I.drawOn(blueprint, 4*inch, 3*inch)
        blueprint.line(0.3*inch,2.89*inch, 8*inch, 2.89*inch)
        #================================ End: Drawing Time Spent BAR Graph =========================================

        #================================ Begin: Drawing Correct Vs Wrong Vs Attempted Vs UnAttempted PIE Chart =========================================
        fig = make_subplots(rows=1, cols=3, specs=[[{'type':'domain'}, {'type':'domain'},{'type':'domain'}]],subplot_titles=("Ateempts vs Not Attempts", "Correct Vs Wrong", "Correct Vs Wrong Vs UnAttempted"))
        fig.add_trace(go.Pie(labels=["Attempted","Not Attempted"], values=attempts_count, hole=.4, name="Ateempts vs Not Attempts"),
                    1, 1)
        fig.add_trace(go.Pie(labels=["Correct","Wrong"], values=[correct_count, wrong_count],hole=.4, name="Correct Vs Wrong"),
                    1, 2)
        fig.add_trace(go.Pie(labels=["Correct","Wrong", "Not Attempted"], values=[correct_count, wrong_count, attempts_count[1]], hole=.4, name="Correct Vs Wrong Vs UnAttempte"),
                    1, 3)
        fig.update_layout(legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
            ),
            font=dict(
                family="Courier",
                size=16,
                color="black"
            )
        
        )
        for annotation in fig['layout']['annotations']: 
            annotation['y']= 0.8
            annotation['font']['size'] = 12
        fig.layout.annotations[0].update(x=0.025)
        fig.layout.annotations[2].update(x=0.98)
        fig.write_image("./graphs/{}_attempted.png".format(arr[0]))
        image_path = "./graphs/{}_attempted.png".format(arr[0])
        I = Image(image_path)
        I.drawHeight = 5.8*inch*I.drawHeight / I.drawWidth
        I.drawWidth = 7.25*inch
        I.drawOn(blueprint, 0.5*inch, -1.28*inch)
        #================================ End: Drawing Correct Vs Wrong Vs Attempted Vs UnAttempted PIE Chart =========================================
        blueprint.save() #Saving the Pdf






