import requests
import json
import datetime
from docxtpl import DocxTemplate
import re
import os
import io
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS # Importa a extensão CORS

app = Flask(__name__)
CORS(app) # Habilita o CORS para todas as rotas

# Configurações do servidor Flask
app.config['DEBUG'] = True
app.config['JSON_AS_ASCII'] = False # Permite caracteres especiais no JSON

# Credenciais e URLs da API (MANTIDAS AQUI NO BACK-END PARA SEGURANÇA)
login_url = 'https://jcrisco.com.br/logtrack/controllers/usuario_login.php'
login_data = {
    'usuario': '55032620858',
    'senha': '280406'
}
consulta_url = 'https://jcrisco.com.br/logtrack/controllers/consultaPesquisas.php'
veiculo_lista_url = 'https://jcrisco.com.br/logtrack/controllers/veiculo_lista.php'
veiculo_edita_url = 'https://jcrisco.com.br/logtrack/controllers/veiculo_edita.php'

# Rota para buscar os dados de uma viagem
@app.route('/api/search', methods=['POST'])
def search_trip():
    data = request.json
    id_pesquisa = data.get('id')

    if not id_pesquisa:
        print("Erro: ID de pesquisa não fornecido na requisição.")
        return jsonify({'success': False, 'message': 'ID de pesquisa não fornecido.'}), 400

    print(f"Tentando login na API LogTrack...")
    session = requests.Session()
    try:
        login_response = session.post(login_url, data=login_data, timeout=10)
        login_response.raise_for_status() # Lança um erro para status de resposta ruins (4xx ou 5xx)
        login_result = login_response.json()
    except (requests.exceptions.RequestException, json.JSONDecodeError) as e:
        print(f"Erro no login ou na resposta da API: {e}")
        return jsonify({'success': False, 'message': 'Erro de login ou na resposta da API.'}), 500

    if not login_result.get('success'):
        print(f"Login falhou: {login_result.get('msg', 'Mensagem de erro desconhecida.')}")
        return jsonify({'success': False, 'message': login_result.get('msg', 'Erro no login.')}), 401

    print("Login bem-sucedido. Consultando dados da viagem...")
    consulta_data = {
        "page": 1,
        "rows": 100,
        "filterRules": json.dumps([
            {"field": "id_bl", "op": "contains", "value": id_pesquisa}
        ])
    }
    
    try:
        consulta_response = session.post(consulta_url, data=consulta_data, timeout=10)
        consulta_response.raise_for_status()
        resultado = consulta_response.json()
    except (requests.exceptions.RequestException, json.JSONDecodeError) as e:
        print(f"Erro na busca da viagem ou na resposta da API: {e}")
        return jsonify({'success': False, 'message': 'Erro na busca ou na resposta da API.'}), 500

    if not resultado.get('rows'):
        print(f"ID '{id_pesquisa}' não encontrado.")
        return jsonify({'success': False, 'message': 'ID não encontrado ou erro na busca.'}), 404

    detalhes_da_viagem = resultado['rows'][0]

    placas_para_buscar = [
        detalhes_da_viagem.get('veiculo'),
        detalhes_da_viagem.get('reboque1'),
        detalhes_da_viagem.get('reboque2'),
        detalhes_da_viagem.get('reboque3')
    ]
    
    dados_veiculos_completos = {}
    print(f"Buscando dados de veículos para as placas: {placas_para_buscar}")
    for placa in placas_para_buscar:
        if placa:
            try:
                consulta_veiculo_response = session.post(f'{veiculo_lista_url}?placa={placa}', timeout=10)
                c_json = consulta_veiculo_response.json()
                if c_json:
                    id_veiculo = c_json[0].get('id')
                    if id_veiculo:
                        url_dados_veiculo = f"{veiculo_edita_url}?id={id_veiculo}"
                        response_dados = session.get(url_dados_veiculo, timeout=10)
                        dados_veiculo = response_dados.json()
                        dados_veiculos_completos[placa] = {
                            'placa': placa,
                            'proprietario': dados_veiculo.get('proprietario', ''),
                            'proprietario_cpfcnpj': dados_veiculo.get('proprietario_cpfcnpj', ''),
                            'renavam': dados_veiculo.get('renavan', ''),
                        }
            except (requests.exceptions.RequestException, json.JSONDecodeError) as e:
                print(f"Erro ao buscar detalhes do veículo {placa}: {e}")

    return jsonify({
        'success': True,
        'trip': detalhes_da_viagem,
        'vehicles': dados_veiculos_completos
    })

# Rota para gerar o documento e iniciar o download
@app.route('/api/generate-document', methods=['POST'])
def generate_document():
    request_data = request.json
    selected_items = request_data.get('items')
    detalhes_da_viagem = request_data.get('trip')
    dados_veiculos_completos = request_data.get('vehicles')

    if not selected_items or not detalhes_da_viagem or not dados_veiculos_completos:
        print("Erro: Dados necessários para gerar o documento não foram fornecidos.")
        return jsonify({'success': False, 'message': 'Dados necessários não foram fornecidos.'}), 400

    selecionou_motorista = 'motorista' in selected_items
    placas_selecionadas = [item.split(':')[1] for item in selected_items if item.startswith('placa:')]
    num_veiculos_selecionados = len(placas_selecionadas)

    caminho_modelo = ""
    
    if selecionou_motorista:
        if num_veiculos_selecionados == 0:
            caminho_modelo = "cartas_timbradas/TIMBRADA - (01) MOTORISTA.docx"
        elif num_veiculos_selecionados == 1:
            caminho_modelo = "cartas_timbradas/TIMBRADA - (01)_MOTORISTA (01)_VEICULO.docx"
        else:
            caminho_modelo = "cartas_timbradas/TIMBRADA - (01)_MOTORISTA (02)_VEICULOS.docx"
    else:
        if num_veiculos_selecionados == 1:
            caminho_modelo = "cartas_timbradas/TIMBRADA - (01) VEICULO.docx"
        elif num_veiculos_selecionados > 1:
            caminho_modelo = "cartas_timbradas/TIMBRADA - (02) VEICULOS.docx"
        else:
            return jsonify({'success': False, 'message': 'Nenhum modelo de documento encontrado.'}), 404
    
    meses_pt = {
        "January": "Janeiro", "February": "Fevereiro", "March": "Março", "April": "Abril",
        "May": "Maio", "June": "Junho", "July": "Julho", "August": "Agosto",
        "September": "Setembro", "October": "Outubro", "November": "Novembro", "December": "Dezembro"
    }
    mes_hoje = datetime.date.today().strftime('%B')
    mes_portugues = meses_pt.get(mes_hoje, mes_hoje)
    
    data_completa = detalhes_da_viagem.get('validade_data', '')
    match = re.search(r'\d{2}/\d{2}/\d{4}', data_completa)
    apenas_data = match.group(0) if match else ''

    data_completa_consulta = detalhes_da_viagem.get('consulta_datahora', '')
    match_consulta = re.search(r'\d{2}/\d{2}/\d{4}', data_completa_consulta)
    apenas_data_consulta = match_consulta.group(0) if match_consulta else ''

    contexto = {
        'solicitacao': detalhes_da_viagem.get('viagensLiberacaoId', ''),
        'status': detalhes_da_viagem.get('status', ''),
        'motorista': detalhes_da_viagem.get('motorista', '') if selecionou_motorista else '',
        'cpf': detalhes_da_viagem.get('motorista_cpf', '') if selecionou_motorista else '',
        'cliente': detalhes_da_viagem.get('cliente', ''),
        'transportadora': detalhes_da_viagem.get('transportadora', ''),
        'tipo': detalhes_da_viagem.get('tipo_motorista', '') if selecionou_motorista else '',
        'placa1': '', 'proprietario1': '', 'cpf_cnpj1': '', 'renavam1': '',
        'placa2': '', 'proprietario2': '', 'cpf_cnpj2': '', 'renavam2': '',
        'placa_reboque2': '', 'proprietario_reboque2': '', 'cpf_cnpj_reboque2': '', 'renavam_reboque2': '',
        'placa_reboque3': '', 'proprietario_reboque3': '', 'cpf_cnpj_reboque3': '', 'renavam_reboque3': '',
    }
    
    for i, placa in enumerate(placas_selecionadas):
        dados_veiculo = dados_veiculos_completos.get(placa, {})
        if i == 0:
            contexto['placa1'] = dados_veiculo.get('placa', '')
            contexto['proprietario1'] = dados_veiculo.get('proprietario', '')
            contexto['cpf_cnpj1'] = dados_veiculo.get('proprietario_cpfcnpj', '')
            contexto['renavam1'] = dados_veiculo.get('renavam', '')
        elif i == 1:
            contexto['placa2'] = dados_veiculo.get('placa', '')
            contexto['proprietario2'] = dados_veiculo.get('proprietario', '')
            contexto['cpf_cnpj2'] = dados_veiculo.get('proprietario_cpfcnpj', '')
            contexto['renavam2'] = dados_veiculo.get('renavam', '')
        elif i == 2:
            contexto['placa_reboque2'] = dados_veiculo.get('placa', '')
            contexto['proprietario_reboque2'] = dados_veiculo.get('proprietario', '')
            contexto['cpf_cnpj_reboque2'] = dados_veiculo.get('proprietario_cpfcnpj', '')
            contexto['renavam_reboque2'] = dados_veiculo.get('renavam', '')
        elif i == 3:
            contexto['placa_reboque3'] = dados_veiculo.get('placa', '')
            contexto['proprietario_reboque3'] = dados_veiculo.get('proprietario', '')
            contexto['cpf_cnpj_reboque3'] = dados_veiculo.get('proprietario_cpfcnpj', '')
            contexto['renavam_reboque3'] = dados_veiculo.get('renavam', '')
    
    contexto.update({
        'data_pesquisa': apenas_data_consulta,
        'validade': apenas_data,
        'dia': str(datetime.date.today().day),
        'mes': mes_portugues,
        'ano': str(datetime.date.today().year)
    })
    
    try:
        # Abre o modelo
        doc = DocxTemplate(caminho_modelo)
        doc.render(contexto)
        
        # Salva o arquivo em um buffer na memória, sem salvar no disco
        docx_buffer = io.BytesIO()
        doc.save(docx_buffer)
        docx_buffer.seek(0)
        
        # Envia o arquivo para o front-end
        filename = f"Relatorio_{detalhes_da_viagem['viagensLiberacaoId']}.docx"
        return send_file(docx_buffer, as_attachment=True, download_name=filename)

    except FileNotFoundError:
        return jsonify({'success': False, 'message': 'O arquivo de modelo não foi encontrado.'}), 404
    except Exception as e:
        return jsonify({'success': False, 'message': f'Erro ao processar o documento: {str(e)}'}), 500

if __name__ == '__main__':
    print("Iniciando o servidor Flask...")
    app.run(port=5000)
