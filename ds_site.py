import os

from flask import Flask, render_template, request, url_for
import openpyxl

frontend = Flask(__name__)

AppRoot = os.path.dirname(os.path.abspath(__file__))

@frontend.route("/cover")
@frontend.route("/")
def cover():
    return render_template('cover.html')

@frontend.route("/upload", methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        PriceListTarget = os.path.join(AppRoot, 'PriceLists/')
        ConfigTarget = os.path.join(AppRoot, 'Config/')
        
        conf_file = ''
        price_file = ''

        if not os.path.isdir(PriceListTarget):
            os.mkdir(PriceListTarget)
            
        if not os.path.isdir(ConfigTarget):
            os.mkdir(ConfigTarget)
            
        for file in request.files.getlist('file'):
            filename = file.filename
            if filename.find('Config') != -1:
                destination = '/'.join([ConfigTarget, filename])
                file.save(destination)
                conf_file = destination
            else:
                destination = '/'.join([PriceListTarget, filename])
                file.save(destination)
                price_file = destination
               
        return startFuf(conf_file, price_file) 
            
    return render_template('Upload_PriceList.html')

def startFuf(conf_file, price_file):
    contract=openpyxl.load_workbook(conf_file, data_only=True)
    names=contract.sheetnames
    return names[0]

if __name__ == '__main__':
    frontend.run(host='127.0.0.1', port=8080,debug=True)
