# Usa uma imagem base do Python 3.9
FROM python:3.9-slim

# Define o diretório de trabalho dentro do contêiner
WORKDIR /app

# Instala o LibreOffice e outras dependências do sistema
RUN apt-get update && apt-get install -y libreoffice

# Copia os arquivos de configuração do Gunicorn e os scripts da API
COPY gunicorn.conf.py ./
COPY run.py ./
COPY api.py ./

# Copia o arquivo de requisitos e instala as dependências
COPY requirements.txt ./
RUN pip install -r requirements.txt

# Copia a pasta com os modelos de documentos
COPY cartas_timbradas ./cartas_timbradas

# Exponha a porta
EXPOSE 10000

# Comando para iniciar o servidor Gunicorn
CMD ["gunicorn", "-c", "gunicorn.conf.py", "run:app"]
