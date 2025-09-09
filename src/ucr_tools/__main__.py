import pandas as pd
from pathlib import Path
from core import utils
import smtplib, ssl
import os, sys
import traceback

def main():
    while True:
        inp=None
        while not inp:
            try:
                inp = int(input('1 - Format list\n2 - Send emails\n> '))
            except ValueError:
                print('please enter a number')
        if inp == 1:
            students_path = Path(input('Input file with student list:\n> ').strip('"'))
            df = pd.read_excel(students_path)
            df = utils.remove_white_rows(df)
            df = utils.split_names(df)
            df = utils.create_zipgrade_id(df)
            df.dropna(axis=1,how='all').to_clipboard(index=False)
            print("Press ctrl-v to paste in excel")
            print('exiting program')
            sys.exit()
        
        elif inp == 2:
            grades_path = input('Input zipgrade report file\n> ').strip('"')
            if grades_path: # Generate df with grades
                grades_path = Path(grades_path)
                grupo = pd.read_csv(grades_path).rename({'Student First Name':'nombre', 'Percent Correct':'nota', 'External Ref':'correo', 'Student ID':'student_id'}, axis=1)
                keep_cols = ['nombre', 'correo', 'nota', 'student_id']
                grupo = grupo[keep_cols]
            else:
                raise ValueError('El directorio no es valido')
                
            pdf_files = input('If pdfs are attached input folder to pdfs\n>').strip('"')
            if pdf_files: # Generate df with pdfs
                pdf_files=Path(pdf_files)
                pdf_files = utils.get_student_pdf(pdf_files)
                grupo = pd.merge(grupo, pdf_files, how='left', on='student_id')
            else:
                grupo = grupo.assign(pdf_dir=None)
                
            try: #Create and send email
                context = ssl.create_default_context()
                server = smtplib.SMTP_SSL('smtp.ucr.ac.cr', 465, context=context)
                user, password = os.getenv('ucr_email_user'), os.getenv('ucr_email_password')
                server.login(user, password)
                
                plantilla = Path(input("Plantilla a usar:\n> ").strip('"'))
                subject = input('Input email subject: ')
                test = input('Input test name: ')
                total = int(input('Input total percentage: '))
                
                with plantilla.open('rt') as file:
                        plantilla = file.read()
                for idx, row in grupo.iterrows():
                    format_args = {'student_name' : str(row['nombre']).split()[0].title(),
                                    'test' : test,
                                    'grade' : row['nota'],
                                    'parcent' : round(row['nota']/100*total,2),
                                    'total' : total,
                                    'sign':user.split('.')[0].title()}
                    
                    msg = utils.base_email_ucr(user=user, to=row['correo'], email_body=plantilla, format_args=format_args,subject=subject, attachment=row['pdf_dir'])
                    server.sendmail(to_addrs=[msg['To'], user], from_addr=user, msg = msg.as_string())
                    print(f'mail sent to > {idx+1}:  {row["nombre"]}')
            except Exception as e:
                traceback.print_exc()
            finally:
                server.quit()
                print('exiting program')
                sys.exit()

if __name__=='__main__':
    main()