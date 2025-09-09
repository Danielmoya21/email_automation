import pandas as pd
from pathlib import Path
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

def remove_white_rows(df:pd.DataFrame) -> pd.DataFrame:
    return df[~(df.isna().sum(axis=1)>4)].reset_index(drop=True)#-- Solo necesario si la lista no ha sido depurada
    
def split_names(df):
    """
    This function assumes that column name is nombre and persons name follows the following order:
    Lastname1 Lastname2 Name
    """
    split_names = df['nombre'].str.split(n=2, expand=True)
    df['nombre'] = split_names[2]
    df['apellido'] = split_names[0] +" "+ split_names[1]
    for c in 'nombre apellido'.split():
        df[c] = df[c].str.title()
    return df

# Create Id without letters
def create_zipgrade_id(df):
    df['id'] = df['carné'].str[1:].str.replace('\D', '0', regex=True).astype(float).round(0)
    return df

def base_email_ucr(user:str, to:str|list, email_body:str, format_args:str, subject:str, attachment:Path):
    msg = MIMEMultipart()
    msg['From'] = user
    msg['To'] = to
    msg['Subject'] = subject
    
    if attachment:
        format_args['attachment'] = 'Le adjunto también la hoja de respuestas con la revisión.'
        with open(attachment, 'rb') as file:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(file.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{attachment.name}"')
            msg.attach(part)
    else:
        format_args['attachment'] = ''
        
    msg.attach(MIMEText(email_body.format(**format_args), 'plain'))
    return msg
    
def get_student_pdf(pdf_files) -> pd.DataFrame:
    student_pdf = []
    for i in pdf_files.iterdir():
        id = int(i.stem.rsplit('-', maxsplit=2)[-2])
        student_pdf.append((id, i.absolute()))
        
    return pd.DataFrame(student_pdf, columns=['student_id', 'pdf_dir'])

