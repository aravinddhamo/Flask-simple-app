from flask_codes import app,db
from flask_codes.models import User
from flask import render_template,url_for,redirect,flash,request,Response
from flask_codes.forms import RegisterForm,SearchForm
import io,xlwt

@app.route('/home',methods=['POST','GET'])
def home():
    form=RegisterForm()
    if form.validate_on_submit():
        user=User(ClientId=form.ClientId.data,ClientName=form.ClientName.data,ClientType=form.ClientType.data,ClientBSI=form.ClientBSI.data,ClientStatus=form.ClientStatus.data)
        db.session.add(user)
        db.session.commit()
        flash(f'Details are added successfully for {form.ClientId.data}',category='success')
        return redirect(url_for('home'))
    return render_template('homepage.html',title='Home Page',form=form)



@app.route('/search',methods=['POST','GET'])
def search():
   
    form=SearchForm()
    if form.validate_on_submit():
        
        user=User.query.filter_by(ClientId=form.ClientId.data).first()
        try:
         if form.ClientId.data==user.ClientId:
           flash(f'Please find the details for Client ID : {form.ClientId.data}',category='success')
           clientids=form.ClientId.data
           clientnames=user.ClientName
           clienttypes=user.ClientType
           clientbsis=user.ClientBSI
           clientstatuss=user.ClientStatus
           #print(clientids)
           #print(clientnames)
           #print(clienttypes)
           #print(clientbsis)
           #print(clientstatuss)
           
    
        
           return redirect(url_for('search_result',clientids=clientids,clientnames=clientnames,clienttypes=clienttypes,clientbsis=clientbsis,clientstatuss=clientstatuss))
        except:
            flash(f'Invalid ID: {form.ClientId.data}.Please search for the valid ID',category='danger')
            return redirect(url_for('search'))
    return render_template('search.html',title='Search Page',form=form)

@app.route('/search_result/<clientids>/<clientnames>/<clienttypes>/<clientbsis>/<clientstatuss>')
def search_result(clientids,clientnames,clienttypes,clientbsis,clientstatuss):
    print(clientids)
    print(clientnames)
    print(clienttypes)
    print(clientbsis)
    print(clientstatuss)
    
    return render_template('search_result.html',title='Result Page',clientids=clientids,clientnames=clientnames,clienttypes=clienttypes,clientbsis=clientbsis,clientstatuss=clientstatuss)

@app.route('/')
def main_page():
    return render_template('main_page.html',title='ELF Page')

@app.route('/all_tskl')
def all_tskl():
    details=User.query.all()
    return render_template('all_tskl.html',title='Tasklist Page',details=details)

#Download task-to download the datas into excel file
@app.route('/download_report')
def download_report():
    details=User.query.all()
    output=io.BytesIO()
    #create Workbook object
    workbook=xlwt.Workbook()
    #add a sheet
    sh=workbook.add_sheet('Client Report')
 
 #Adding first row as headers
    sh.write(0,0,'ClientId')
    sh.write(0,1,'ClientName')
    sh.write(0,2,'ClientType')
    sh.write(0,3,'ClientBSI')
    sh.write(0,4,'ClientStatus')
    
    #inserting every values inot excel
    idx=0
    for rows in details:
        sh.write(idx+1,0,rows.ClientId)
        sh.write(idx+1,1,rows.ClientName)
        sh.write(idx+1,2,rows.ClientType)
        sh.write(idx+1,3,rows.ClientBSI)
        sh.write(idx+1,4,rows.ClientStatus)
        idx+=1

    workbook.save(output)
    output.seek(0)

     
    return Response(output,mimetype='application/ms-excel',headers={"Content-Disposition":"attachment;filename=Client_report.xls"})