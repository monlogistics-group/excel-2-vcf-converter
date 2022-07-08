from flask import Flask, render_template, request, send_file
import pandas as pd
import vobject
app = Flask(__name__)
@app.route('/', methods=['GET', 'POST'])
def index():
    return render_template('index.html')
@app.route('/data', methods=['GET', 'POST'])
def data():
    if request.method == "POST":
        file = request.files['upload-file']
        data = pd.read_excel(file)  
        data.columns= [['№', 'Овог', 'Нэр', 'Албан тушаал', 'Компани', 'Алба', 'Хэлтэс', 'Байршил', 'Утасны дугаар 1', 'Утасны дугаар 2', 'Дотуур холбоо', 'Э-Мэйл хаяг']]
        del data['Алба']
        del data['Байршил']
        del data['№']
        del data['Дотуур холбоо']
        data=data.drop(data.index[0])
        for index,rows in data.iterrows():
            if(pd.isna(rows['Нэр'])):
                data = data.drop([index])
        data.reset_index(drop=True, inplace=True)
     
        contact_demo = []
        for i in range(0, len(data['Нэр'])):
           
         
            vCard = vobject.vCard()
            vCard.add('n')
            vCard.n.value = vobject.vcard.Name( family=data.loc[i]['Овог'], given=data.loc[i]['Нэр'] )
            vCard.add('Fn')
            vCard.fn.value = data.loc[i]['Овог']

            vCard.add('TEL')
            vCard.tel.value = str(data.loc[i]['Утасны дугаар 1'])
            vCard.tel.type_param = 'HOME'
            vCard.add('TITLE')
            vCard.title.value = str(data.loc[i]['Албан тушаал'])
            vCard.add('NOTE')
            vCard.note.value =str(data.loc[i]["Компани"]) + " " + str(data.loc[i]['Хэлтэс'])  
            vCard.add('EMAIL')
            vCard.email.value = str(data.loc[i]['Э-Мэйл хаяг'])
            vCard.email.type_param = 'INTERNET'
            contact_demo.append(vCard.serialize())
    

        with open('fname.vcf', "w", encoding="utf-8") as f:
            for line in contact_demo:
                f.write(line+'\n')
                f.closed
        data.to_excel('data.xlsx')
        return send_file('fname.vcf')


if __name__=='__main__':
    app.run(debug=True)