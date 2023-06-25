import openpyxl
import pandas as pd
import matplotlib.pyplot as plt
from flask import Flask, render_template, request, redirect

app = Flask(__name__)

data = []

@app.route('/')
def index():
    return render_template('index.html', data=data)

@app.route('/add', methods=['GET', 'POST'])
def add():
    if request.method == 'POST':
        name = request.form['name']
        age = request.form['age']
        data.append({'Name': name, 'Age': age})
        return redirect('/')
    return render_template('add.html')

@app.route('/edit/<int:index>', methods=['GET', 'POST'])
def edit(index):
    if request.method == 'POST':
        name = request.form['name']
        age = request.form['age']
        data[index] = {'Name': name, 'Age': age}
        return redirect('/')
    return render_template('edit.html', index=index, person=data[index])

@app.route('/delete/<int:index>', methods=['POST'])
def delete(index):
    del data[index]
    return redirect('/')

@app.route('/upload', methods=['GET', 'POST'])
def upload():
    if request.method == 'POST':
        file = request.files['file']
        if file.filename.endswith('.xlsx'):
            wb = openpyxl.load_workbook(file)
            sheet = wb.active
            data.clear()
            for row in sheet.iter_rows(values_only=True):
                name, age = row
                data.append({'Name': name, 'Age': age})
            return redirect('/')
    return render_template('upload.html')

@app.route('/visualize')
def visualize():
    df = pd.DataFrame(data)
    df = df[df['Age'].apply(lambda x: str(x).isnumeric())] 
    df['Age'] = df['Age'].astype(int)  
    plt.bar(df['Name'], df['Age'])
    plt.xlabel('Name')
    plt.ylabel('Age')
    plt.title('Age Distribution')
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig('static/plot.png')
    return render_template('visualize.html')

if __name__ == '__main__':
    app.run(debug=True)