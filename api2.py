# -*- coding: utf-8 -*-
from flask import Flask, request, jsonify, send_file
import requests
import json
import datetime
from docxtpl import DocxTemplate
import re
import os
from flask_cors import CORS
import subprocess
import shutil

app = Flask(__name__)
CORS(app)

# O Render irá criar este diretório se não existir.
# home_dir = os.path.expanduser('~')
# pasta_de_saida = os.path.join(home_dir, "Downloads", "cartas timbradas feitas")
template_dir = 'carta_timbrada'

def formatar_retorno(sucesso, msg):
    return {'success': sucesso, 'message': msg}

def buscar_dados_viagem(id_pesquisa):
    session = requests.Session()
    login_url = 'https://jcrisco.com.br/logtrack/controllers/usuario_login.php'

    # Lê as credenciais das variáveis de ambiente
    usuario = os.environ.get('LOGTRACK_USUARIO')
    senha = os.environ.get('LOGTRACK_SENHA')

    if not usuario or not senha:
        return None, "❌ Credenciais de login não configuradas nas variáveis de ambiente."

    login_data = {'usuario': usuario, 'senha': senha}

    try:
        login_response = session.post(login_url, data=login_data)
        login_result = login_response.json()
    except Exception as e:
        print(f"Erro ao tentar login ou resposta inválida: {e}")
        return None, "❌ Erro ao tentar login ou resposta inválida."

    if not login_result.get('success'):
        print(f"Erro no login: {login_result.get('msg', 'Sem mensagem de erro')}")
        return None, "❌ Erro no login: " + login_result.get('msg', 'Sem mensagem de erro')

    consulta_url = 'https://jcrisco.com.br/logtrack/controllers/consultaPesquisas.php'
    consulta_data = {
        "page": 1,
        "rows": 100,
        "filterRules": json.dumps([{"field": "id_bl", "op": "contains", "value": id_pesquisa}])
    }

    try:
        consulta_response = session.post(consulta_url, data=consulta_data)
        resultado = consulta_response.json()
    except Exception as e:
        print(f"Erro ao consultar a viagem ou resposta inválida: {e}")
        return None, "❌ Erro ao consultar a viagem ou resposta inválida."

    if not resultado.get('rows'):
        return None, "⚠️ Nenhum dado de viagem encontrado para o ID informado."

    detalhes_da_viagem = resultado['rows'][0]

    placas_para_buscar = []
    for campo in ['veiculo','reboque1','reboque2','reboque3']:
        if detalhes_da_viagem.get(campo):
            placas_para_buscar.append(detalhes_da_viagem.get(campo))

    dados_veiculos_completos = {}
    for placa in placas_para_buscar:
        try:
            consulta_veiculo_url = f'https://jcrisco.com.br/logtrack/controllers/veiculo_lista.php?placa={placa}'
            c_json = session.post(consulta_veiculo_url).json()
            if c_json and c_json[0].get('id'):
                id_veiculo = c_json[0]['id']
                url_dados_veiculo = f"https://jcrisco.com.br/logtrack/controllers/veiculo_edita.php?id={id_veiculo}&_="
                dados_veiculo = session.get(url_dados_veiculo).json()
                dados_veiculos_completos[placa] = {
                    'placa': dados_veiculo.get('placa', ''),
                    'proprietario': dados_veiculo.get('proprietario', ''),
                    'proprietario_cpfcnpj': dados_veiculo.get('proprietario_cpfcnpj', ''),
                    'renavam': dados_veiculo.get('renavan', ''),
                }
        except Exception as e:
            print(f"Erro ao buscar dados para a placa {placa}: {e}")
            pass

    return {'trip_data': detalhes_da_viagem, 'vehicle_data': dados_veiculos_completos}, None

@app.route('/search_data', methods=['POST'])
def search_data():
    data = request.get_json()
    id_pesquisa = data.get('id_pesquisa')

    if not id_pesquisa:
        return jsonify(formatar_retorno(False, "O parâmetro 'id_pesquisa' é obrigatório.")), 400
    
    dados, erro = buscar_dados_viagem(id_pesquisa)
    if erro:
        return jsonify(formatar_retorno(False, erro)), 500
    
    return jsonify({
        'success': True,
        'message': 'Dados de viagem encontrados.',
        'trip_data': dados['trip_data'],
        'vehicle_data': dados['vehicle_data']
    })

def docx_para_pdf(caminho_docx, caminho_pdf):
    """
    Converte um arquivo DOCX para PDF usando o Pandoc.
    Pandoc precisa estar instalado no ambiente do servidor.
    """
    try:
        subprocess.run(['pandoc', '-o', caminho_pdf, caminho_docx], check=True)
    except FileNotFoundError:
        raise RuntimeError("Pandoc não está instalado ou não está no PATH. Instale-o no seu ambiente de deploy.")
    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"Erro ao converter DOCX para PDF com Pandoc: {e}")

@app.route('/generate_docs', methods=['POST'])
def generate_docs():
    data = request.get_json()
    id_pesquisa = data.get('id_pesquisa')
    escolha_str = data.get('escolha_str', '')

    if not id_pesquisa:
        return jsonify(formatar_retorno(False, "O parâmetro 'id_pesquisa' é obrigatório.")), 400
    
    dados, erro = buscar_dados_viagem(id_pesquisa)
    if erro:
        return jsonify(formatar_retorno(False, erro)), 500
    
    detalhes_da_viagem = dados['trip_data']
    dados_veiculos_completos = dados['vehicle_data']

    itens_disponiveis = []
    if detalhes_da_viagem.get('motorista'):
        itens_disponiveis.append("Motorista")

    placas_para_buscar = [detalhes_da_viagem.get(c) for c in ['veiculo','reboque1','reboque2','reboque3'] if detalhes_da_viagem.get(c)]
    placas_disponiveis = [p for p in placas_para_buscar if p in dados_veiculos_completos]
    for p in placas_disponiveis:
        itens_disponiveis.append(f"Placa: {p}")

    if not itens_disponiveis:
        return jsonify(formatar_retorno(False, "⚠️ Nenhum item (motorista ou veículo) encontrado na pesquisa.")), 404

    selecionou_motorista = False
    placas_selecionadas = []

    if not escolha_str or escolha_str == '0':
        selecionou_motorista = bool(detalhes_da_viagem.get('motorista'))
        placas_selecionadas = placas_disponiveis
    else:
        try:
            escolhas = [int(n.strip()) for n in escolha_str.split(',')]
            for escolha in escolhas:
                if 1 <= escolha <= len(itens_disponiveis):
                    item = itens_disponiveis[escolha-1]
                    if item == "Motorista":
                        selecionou_motorista = True
                    elif item.startswith("Placa:"):
                        placas_selecionadas.append(item.split(': ')[1])
        except Exception as e:
            print(f"Erro na entrada de escolha: {e}")
            return jsonify(formatar_retorno(False, "❌ Entrada inválida. Digite números separados por vírgula.")), 400

    if not selecionou_motorista and not placas_selecionadas:
        return jsonify(formatar_retorno(False, "⚠️ Nenhum item foi selecionado para o documento.")), 400

    # Seleção do modelo
    caminho_modelo = ""
    n_veiculos = len(placas_selecionadas)
    if selecionou_motorista:
        if n_veiculos == 0:
            caminho_modelo = "TIMBRADA - (01) MOTORISTA.docx"
        elif n_veiculos == 1:
            caminho_modelo = "TIMBRADA - (01)_MOTORISTA (01)_VEICULO.docx"
        else:
            caminho_modelo = "TIMBRADA - (01)_MOTORISTA (02)_VEICULOS.docx"
    else:
        if n_veiculos == 1:
            caminho_modelo = "TIMBRADA - (01) VEICULO.docx"
        elif n_veiculos > 1:
            caminho_modelo = "TIMBRADA - (02) VEICULOS.docx"
            
    caminho_completo_modelo = os.path.join(template_dir, caminho_modelo)

    if not os.path.exists(caminho_completo_modelo):
        print(f"Modelo '{caminho_completo_modelo}' não encontrado.")
        return jsonify(formatar_retorno(False,f"Modelo '{caminho_modelo}' não encontrado.")),500

    # Preparar contexto
    meses_pt = {
        "January": "Janeiro", "February": "Fevereiro", "March": "Março", "April": "Abril",
        "May": "Maio", "June": "Junho", "July": "Julho", "August": "Agosto",
        "September": "Setembro", "October": "Outubro", "November": "Novembro", "December": "Dezembro"
    }
    mes_portugues = meses_pt.get(datetime.date.today().strftime('%B'), datetime.date.today().strftime('%B'))

    contexto = {
        'solicitacao': str(detalhes_da_viagem.get('viagensLiberacaoId','')),
        'status': detalhes_da_viagem.get('status',''),
        'motorista': detalhes_da_viagem.get('motorista','') if selecionou_motorista else '',
        'cpf': detalhes_da_viagem.get('motorista_cpf','') if selecionou_motorista else '',
        'cliente': detalhes_da_viagem.get('cliente',''),
        'transportadora': detalhes_da_viagem.get('transportadora',''),
        'tipo': detalhes_da_viagem.get('tipo_motorista','') if selecionou_motorista else '',
        'placa1':'','proprietario1':'','cpf_cnpj1':'','renavam1':'',
        'placa2':'','proprietario2':'','cpf_cnpj2':'','renavam2':'',
        'placa_reboque2':'','proprietario_reboque2':'','cpf_cnpj_reboque2':'','renavam_reboque2':'',
        'placa_reboque3':'','proprietario_reboque3':'','cpf_cnpj_reboque3':'','renavam_reboque3':'',
        'data_pesquisa':'','validade':'','dia':str(datetime.date.today().day),
        'mes':mes_portugues,'ano':str(datetime.date.today().year)
    }

    for i, placa in enumerate(placas_selecionadas):
        v = dados_veiculos_completos.get(placa,{})
        if i==0:
            contexto.update({'placa1':v.get('placa'),'proprietario1':v.get('proprietario'),'cpf_cnpj1':v.get('proprietario_cpfcnpj'),'renavam1':v.get('renavam')})
        elif i==1:
            contexto.update({'placa2':v.get('placa'),'proprietario2':v.get('proprietario'),'cpf_cnpj2':v.get('proprietario_cpfcnpj'),'renavam2':v.get('renavam')})
        elif i==2:
            contexto.update({'placa_reboque2':v.get('placa'),'proprietario_reboque2':v.get('proprietario'),'cpf_cnpj_reboque2':v.get('proprietario_cpfcnpj'),'renavam_reboque2':v.get('renavam')})
        elif i==3:
            contexto.update({'placa_reboque3':v.get('placa'),'proprietario_reboque3':v.get('proprietario'),'cpf_cnpj_reboque3':v.get('proprietario_cpfcnpj'),'renavam_reboque3':v.get('renavam')})

    # Datas
    validade = detalhes_da_viagem.get('validade_data','')
    consulta_data = detalhes_da_viagem.get('consulta_datahora','')
    contexto['validade'] = re.search(r'\d{2}/\d{2}/\d{4}', validade).group(0) if validade else ''
    contexto['data_pesquisa'] = re.search(r'\d{2}/\d{2}/\d{4}', consulta_data).group(0) if consulta_data else ''

    # Cria um diretório temporário para os arquivos
    temp_dir = 'temp_docs'
    os.makedirs(temp_dir, exist_ok=True)

    # Nome do arquivo
    nome_base = f"Ficha de liberacao {id_pesquisa}"
    if n_veiculos > 0 and n_veiculos < len(placas_disponiveis):
        nome_base += "_" + "_".join(placas_selecionadas)

    caminho_docx_temporario = os.path.join(temp_dir, f"{nome_base}.docx")
    caminho_pdf_temporario = os.path.join(temp_dir, f"{nome_base}.pdf")

    try:
        doc = DocxTemplate(caminho_completo_modelo)
        doc.render(contexto)
        doc.save(caminho_docx_temporario)

        # Converte para PDF
        docx_para_pdf(caminho_docx_temporario, caminho_pdf_temporario)
        
        # Envia o arquivo e depois o remove para liberar espaço
        response = send_file(caminho_pdf_temporario, as_attachment=True)
        return response

    except Exception as e:
        print(f"❌ Erro na geração do documento: {e}")
        return jsonify(formatar_retorno(False,f"❌ Ocorreu um erro ao processar o documento: {e}")),500
    finally:
        # Limpa o diretório temporário
        shutil.rmtree(temp_dir, ignore_errors=True)

if __name__ == '__main__':
    app.run(debug=True)
