from flask import Flask,request,render_template,send_file
import grouplib
app = Flask(__name__)

@app.get("/")
def hello_world():
    return render_template('index.html',stat='')
@app.post("/group")
def response():
    subjectCode=request.form.get('subjectCode')
    classCode=request.form.get('classCode')
    groupSize=int(request.form.get('groupSize'))
    print(subjectCode,classCode,groupSize)
    final_result=grouplib.webSearch(classCode,subjectCode,groupSize)
    if len(final_result)<1:
        return render_template('index.html',stat="找不到對應課程資訊")
    grouplib.logExcel(final_result)
    return send_file('./result.xlsx')


if __name__=='__main__':
    app.run(host='0.0.0.0',port=5555,debug=True)