# -*- coding: utf-8 -*-
import os
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from datetime import datetime
import re
import tempfile


class ExportManager:
    def __init__(self, database):
        self.database = database

    def exportar_excel(self, caminho_arquivo, filtros=None):
        try:
            servicos, _ = self.database.listar_servicos(filtros=filtros, itens_por_pagina=None)  # Exporta todos
            if not servicos:
                return False

            df = pd.DataFrame(servicos)

            colunas = {
                'id': 'Protocolo', 'data_solicitacao': 'Data Solicitação', 'cpf': 'CPF', 'nome': 'Nome',
                'inscricao_municipal': 'Inscrição Municipal', 'telefone': 'Telefone', 'bairro': 'Bairro',
                'rua': 'Rua', 'numero': 'Número', 'referencia': 'Referência', 'quadra': 'Quadra',
                'lote': 'Lote', 'numero_fossas': 'Número de Fossas', 'status': 'Status',
                'data_chegada': 'Hora Chegada', 'data_saida': 'Hora Saída', 'data_conclusao': 'Data Conclusão',
                'placa_veiculo': 'Placa Veículo', 'motorista': 'Motorista', 'ajudante': 'Ajudante',
                'observacao_empresa': 'Observações'
            }

            df = df.rename(columns=colunas)
            df.to_excel(caminho_arquivo, index=False, sheet_name='Serviços')
            return True
        except Exception as e:
            print(f"Erro ao exportar para Excel: {e}")
            return False

    def _footer_callback(self, canvas, doc):
        canvas.saveState()
        y = 1.0 * cm
        center_x = doc.pagesize[0] / 2
        canvas.setFont('Helvetica', 12)
        canvas.drawCentredString(center_x, y + 3 * cm, "Atesto a conclusão do serviço:")
        canvas.drawCentredString(center_x, y + 2 * cm, "___________________________________")
        canvas.drawCentredString(center_x, y + 1.3 * cm, "Solicitante")
        canvas.setFont('Helvetica', 8)
        canvas.drawCentredString(center_x, y + 0.5 * cm, "[NOME DA EMPRESA]")
        canvas.drawCentredString(center_x, y + 0.2 * cm,
                                 "Rua Florida, Minha Cidade - Meu Estado")
        canvas.drawCentredString(center_x, y - 0.1 * cm, "Email: meuemail@mail.com")
        canvas.restoreState()

    def gerar_pdf(self, id_servico, caminho_arquivo):
        try:
            servico = self.database.obter_servico(id_servico)
            if not servico:
                return False

            doc = SimpleDocTemplate(caminho_arquivo, pagesize=A4, rightMargin=2 * cm, leftMargin=2 * cm,
                                    topMargin=2 * cm, bottomMargin=5 * cm)

            styles = getSampleStyleSheet()
            styles.add(
                ParagraphStyle(name='TituloPrincipal', parent=styles['Heading1'], alignment=TA_CENTER, fontSize=15,
                               spaceAfter=12))
            styles.add(ParagraphStyle(name='Subtitulo', parent=styles['Heading2'], alignment=TA_CENTER, fontSize=14,
                                      spaceAfter=6))
            styles.add(
                ParagraphStyle(name='Inform', fontName='Helvetica', fontSize=13, leading=14, alignment=TA_LEFT))
            styles.add(ParagraphStyle(name='CheckInfo', parent=styles['Inform'], leftIndent=1 * cm))

            styles.add(
                ParagraphStyle(name='NormalFix', fontName='Helvetica', fontSize=11, leading=14, alignment=TA_LEFT))
            styles.add(ParagraphStyle(name='Checkbox', parent=styles['NormalFix'], leftIndent=1 * cm))

            # Máscara para o CPF (ex: 123.***.***-34)
            cpf_limpo = re.sub(r'\D', '', servico['cpf'])
            cpf_mascarado = f"{cpf_limpo[:3]}.***.***-{cpf_limpo[-2:]}"

            elements = []
            elements.append(
                Paragraph(f"SOLICITAÇÃO DE SERVIÇO DE VACOL - Protocolo nº {servico['id']}", styles['TituloPrincipal']))
            data_solicitacao = datetime.strptime(servico['data_solicitacao'], '%Y-%m-%d %H:%M:%S').strftime('%d/%m/%Y')
            elements.append(Paragraph(f"Serviço Solicitado em {data_solicitacao}", styles['Subtitulo']))
            elements.append(Spacer(1, 0.5 * cm))

            elements.append(Paragraph(f"<b>CPF:</b> {cpf_mascarado}", styles['Inform']))
            elements.append(Paragraph(f"<b>Nome:</b> {servico['nome']}", styles['Inform']))
            elements.append(Paragraph(f"<b>Tel/Cel:</b> {servico['telefone']}", styles['Inform']))
            elements.append(Paragraph(f"<b>Bairro:</b> {servico['bairro']}", styles['Inform']))
            elements.append(
                Paragraph(f"<b>Endereço:</b> {servico['rua']}, Nº: {servico['numero']}", styles['Inform']))
            elements.append(Paragraph(f"<b>QD:</b> {servico['quadra'] or ''}   <b>LT:</b> {servico['lote'] or ''}",
                                      styles['Inform']))
            elements.append(Paragraph(f"<b>Referência:</b> {servico['referencia'] or ''}", styles['Inform']))
            elements.append(Paragraph(f"<b>Nº de Fossas:</b> {servico['numero_fossas'] or ''}", styles['Inform']))
            elements.append(Spacer(1, 0.8 * cm))

            chegada = servico.get('data_chegada') or '____:____'
            saida = servico.get('data_saida') or '____:____'
            conclusao = servico.get('data_conclusao') or '__/__/____'
            elements.append(Paragraph("<b>Execução de Serviço</b>", styles['Subtitulo']))
            elements.append(Spacer(1, 0.5 * cm))
            elements.append(Paragraph(
                f"<b>Chegada:</b> {chegada}     <b>Saída:</b> {saida}     <b>(Concluída em:</b> {conclusao}<b>)</b>",
                styles['NormalFix']))
            elements.append(Spacer(1, 0.5 * cm))
            elements.append(Paragraph("<b>Dados da Empresa</b>", styles['Subtitulo']))
            elements.append(Spacer(1, 0.8 * cm))
            elements.append(Paragraph(
                f"<b>Placa:</b> {servico.get('placa_veiculo') or '_____________'}   <b>Motorista:</b> {servico.get('motorista') or '_____________________'}   <b>Ajudante:</b> {servico.get('ajudante') or '_________________'}",
                styles['NormalFix']))
            elements.append(Spacer(1, 1 * cm))
            obs = (servico.get('observacao_empresa') or (f'____________________________________________________________________________\n' +
                   f'____________________________________________________________________________' * 4))
            elements.append(Paragraph(f"<b>Observação:</b> {obs}", styles['NormalFix']))
            elements.append(Spacer(1, 1 * cm))


            # --- SEÇÃO COM CHECKBOXES ---
            elements.append(Paragraph("<b>Situação da Residência (TESTE):</b>", styles['Subtitulo']))
            elements.append(Spacer(1, 0.5 * cm))
            # Altere o texto após os quadrados conforme necessário
            elements.append(Paragraph("(_) Situação 1", styles['NormalFix']))
            elements.append(Paragraph("(_) Situação 2", styles['NormalFix']))
            elements.append(Paragraph("(_) Situação 3", styles['NormalFix']))
            elements.append(Paragraph("(_) Outro Situação 4: ________________________________", styles['NormalFix']))
            elements.append(Spacer(1, 0.5 * cm))


            doc.build(elements, onFirstPage=self._footer_callback, onLaterPages=self._footer_callback)
            return True
        except Exception as e:
            print(f"Erro inesperado durante a geração do PDF: {e}")
            return False

    def visualizar_pdf(self, id_servico):
        try:
            fd, caminho_temporario = tempfile.mkstemp(suffix='.pdf')
            os.close(fd)

            if not self.gerar_pdf(id_servico, caminho_temporario):
                return False

            if os.name == 'nt':  # Windows
                os.startfile(caminho_temporario)
            else:  # Linux, macOS
                os.system(f'xdg-open "{caminho_temporario}"')

            return True
        except Exception as e:
            print(f"Erro ao visualizar PDF: {e}")
            return False