import boto3
import pandas as pd

# Nombre del bucket de Amazon S3
bucket_name = 'abretesesamo'
aws_access_key_id = 'AKIAR5DK3U4IKLOASTL4'
aws_secret_access_key = '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'
region_name = 'us-east-1'  # Ajusta la región según tu configuración
# Crear cliente de Amazon Textract

# Nombre del archivo en el bucket
#document_name = 'DECLARACION PRODUCCION - FACTURA TUNXHO SAS - 05 DE MAYO DE 2023.pdf'


# Nombre del archivo local que deseas cargar
documentName = r"1.jpg"

# Nombre del archivo en el bucket de S3
s3_file_name = 'documento3.jpg'

# Nombre del bucket de S3
bucket_name = 'abretesesamo'

# Crear una instancia del cliente de S3
s3_client = boto3.client('s3',
                          aws_access_key_id=aws_access_key_id,
                          aws_secret_access_key=aws_secret_access_key,
                          region_name=region_name)

# Cargar el archivo en el bucket
s3_client.upload_file(documentName, bucket_name, s3_file_name)

print("Archivo cargado exitosamente en el bucket de S3.")

# Credenciales de AWS

textract = boto3.client('textract', region_name=region_name, aws_access_key_id=aws_access_key_id, aws_secret_access_key=aws_secret_access_key)

# Iniciar el análisis del documento con formularios
response = textract.start_document_analysis(
    DocumentLocation={
        'S3Object': {
            'Bucket': bucket_name,
            'Name': s3_file_name
        }
    },
    FeatureTypes=['FORMS']
)

JobId=response['JobId']

response = textract.get_document_analysis(JobId=JobId)
while response ['JobStatus'] != 'SUCCEEDED':
    response = textract.get_document_analysis(JobId=JobId)
    print('procesing...')

eliminar = s3_client.delete_object(Bucket=bucket_name, Key=s3_file_name)

fields = []
rows = []
text= []
word=[]
for block in response['Blocks']:
    if block['BlockType'] == 'WORD':
        word.append((block['Id'],block['Text']))
    dictionary = {word[i][0]: word[i][1] for i in range(0, len(word), 1)}
    if block['BlockType'] == 'KEY_VALUE_SET':
        for entity in block['EntityTypes']:
            if entity == 'KEY':
                try:
                    nuevo_vector = [dictionary.get(key, key) for key in block['Relationships'][1]['Ids']]
                    fields.append((nuevo_vector,block['Relationships'][0]['Ids'],block['Page']))
                except:
                    pass
            elif entity == 'VALUE':
                try:
                    nuevo_vector = [dictionary.get(key, key) for key in block['Relationships'][0]['Ids']]
                    rows.append((block['Id'],nuevo_vector,block['Page']))
                except:
                    pass
df1 = pd.DataFrame(fields, columns=['key','Id_Value','PageA'])
df2 = pd.DataFrame(rows, columns=['Id_Value','key','PageB'])

for i in range(len(df1)):
    df1.loc[i, 'Id_Value'] = df1['Id_Value'][i][0]


merged_df = pd.merge(df1, df2, on='Id_Value', how='left')
merged_df = merged_df[['key_x','key_y','PageA']]
print(merged_df)