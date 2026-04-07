import os
import shutil
from datetime import datetime, timedelta
import paramiko

# ===== DATA DE ONTEM =====
data_base = datetime.now() - timedelta(days=1)

# formatos diferentes
data_ontem = (datetime.now() - timedelta(days=1)).strftime("%Y%m%d")
data_pasta = data_base.strftime("%d.%m")   # 23.03.2026
data_arquivo = data_base.strftime("%Y%m%d")   # 23032026

# ===== CRIAR PASTA =====
pasta_base = r"Z:\Davi\Via varejo\Tempos\Arquivar\ABRIL"
pasta_destino = os.path.join(pasta_base, data_pasta) 
os.makedirs(pasta_destino, exist_ok=True)

# ===== PASTA DE ORIGEM =====
pasta_origem = r"Z:\Davi\TEMPOS 1"

# ===== PEGAR ARQUIVO =====
arquivos = os.listdir(pasta_origem)
arquivos = [f for f in arquivos if os.path.isfile(os.path.join(pasta_origem, f))]

if len(arquivos) == 0:
    print("Nenhum arquivo encontrado!")
    exit()

if len(arquivos) > 1:
    print("Tem mais de um arquivo na pasta! Deixe apenas um.")
    exit()

arquivo = arquivos[0]
origem = os.path.join(pasta_origem, arquivo)

# =====  NOVO NOME =====
extensao = os.path.splitext(arquivo)[1]
novo_nome = f"Arquivo_Tempos_{data_arquivo}{extensao}"

# ===== DESTINO FINAL =====
destino = os.path.join(pasta_destino, novo_nome)

# =====  MOVER + RENOMEAR =====
shutil.move(origem, destino)

print("Tudo certo! Pasta criada e arquivo organizado ")


# ===== CONFIG SFTP =====
host = "..."   
porta =                          
usuario = "..."
chave_caminho = r"..."

# ===== CONECTAR =====

chave = paramiko.RSAKey.from_private_key_file(chave_caminho)

transport = paramiko.Transport((host, porta))
transport.connect(username=usuario, pkey=chave)
transport.set_keepalive(30)
sftp = paramiko.SFTPClient.from_transport(transport)
print(f"Conectado: {host}")

# ===== CAMINHO REMOTO =====
pasta_remota = "/ARQUIVO_DE_TEMPOS/2024/09.Setembro"
caminho_remoto = f"{pasta_remota}/{novo_nome}"

# ===== CRIAR PASTA NO SFTP =====
try:
    sftp.mkdir(pasta_remota)
except:
    pass  # pasta já existe

# ===== ENVIAR ARQUIVO =====
sftp.put(destino, caminho_remoto)

print("Arquivo enviado para o SFTP")

# ===== FECHAR CONEXÃO =====
sftp.close()
transport.close()