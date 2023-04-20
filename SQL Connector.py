import numpy as np
import pandas as pd
import pymssql as sql
import time
import openpyxl
import warnings

warnings.filterwarnings("ignore")

# Importa o arquivo excel para gerar o dataframe
df_devices = pd.read_excel("Nomedoarquivo.xlsx")
df_devices.head(30)
df_devices.replace(np.nan, None, inplace=True)

# inicia a conexão com SQL Server passando os parâmetros abaixo.
conexao = sql.connect('SERVIDOR', 'USUÁRIO', 'SENHA', 'DATABASE')

# Criando um cursor e executando um LOOP no dataframe para fazer o INSERT no banco SQL Server.
cursor = conexao.cursor()

# Marca o tempo de inicío do insert
inicio = time.time()

# For que vai pegar as colunas e valores e passar as linhas conforme o dataframe lido do arquivo excel
for index, row in df_devices.iterrows():
    sql = "INSERT INTO DefinedDevice(AgencyId, CustomData, Description, DeviceAlias, DeviceId, DeviceSubtypeCode,  \
          DeviceTypeCode, EmployeeId, GpsProtocol, IsTrackable, IsTrackingDeviceForOwner, Owner, UnitId, VehicleId) \
          VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"
    val = (row['AgencyId'], row['CustomData'], row['Description'], row['DeviceAlias'], row['DeviceId'],
           row['DeviceSubtypeCode'], row['DeviceTypeCode'], row['EmployeeId'], row['GpsProtocol'],
           row['IsTrackable'], row['IsTrackingDeviceForOwner'], row['Owner'], row['UnitId'], row['VehicleId'])
    cursor.execute(sql, val)
    conexao.commit()

# Marca o tempo final do insert
final = time.time()

print("Dados inseridos com sucesso no SQL")
print('Tempo de Processamento:', int(final - inicio), 'segundos')

# Select na tabela
tblDevices = pd.read_sql_query('SELECT * FROM DefinedDevices', conexao)

# Fecha a conexão
conexao.close()
