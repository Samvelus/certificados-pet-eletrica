import streamlit as st
import pandas as pd
import io
import zipfile
import os
from datetime import datetime
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Paragraph, Frame
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from PyPDF2 import PdfReader, PdfWriter, PageObject

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Gerador de Certificados PET", layout="wide")

# --- FUNÇÕES UTILITÁRIAS ---

def register_fonts():
    """Tenta registrar fontes, mas usa Helvetica (padrão) se falhar."""
    # Aqui você poderia adicionar opção de upload de fonte futuramente
    # Por padrão, vamos usar Helvetica para garantir portabilidade
    return "Helvetica"

def parse_date_br(date_val):
    if isinstance(date_val, datetime):
        return date_val
    try:
        return pd.to_datetime(date_val, dayfirst=True)
    except:
        return None

def create_overlay_page1(text_line1, date_today, font_name="Helvetica"):
    packet = io.BytesIO()
    # A4 paisagem: width=841.89, height=595.27
    width, height = A4[1], A4[0] 
    c = canvas.Canvas(packet, pagesize=(width, height))

    # Estilo do Parágrafo
    styles = getSampleStyleSheet()
    style = ParagraphStyle(
        'CertStyle',
        parent=styles['Normal'],
        fontName=font_name,
        fontSize=14,
        leading=20,
        alignment=TA_CENTER,
        allowHTML=True
    )

    # Frame centralizado
    frame_width = width * 0.7
    frame_height = 140
    x = (width - frame_width) / 2
    y = height - 410 # Ajuste conforme necessário baseando-se no original

    frame = Frame(x, y, frame_width, frame_height, showBoundary=0)
    p = Paragraph(text_line1, style)
    frame.addFromList([p], c)

    # Data e Local
    c.setFont(font_name, 12)
    c.drawCentredString(width / 2, 60, f"Cuiabá-MT, {date_today}")

    c.save()
    packet.seek(0)
    return packet

def create_overlay_page2(program_items, ministrantes, volume, cert_num, date_today, font_name="Helvetica"):
    packet = io.BytesIO()
    width, height = A4[1], A4[0]
    c = canvas.Canvas(packet, pagesize=(width, height))
    
    margin = 40 * mm
    left_x = margin
    right_x = width/2 + 53*mm
    top_y = height - 45*mm

    # Conteúdo Programático
    styles = getSampleStyleSheet()
    style = ParagraphStyle(
        'ProgStyle',
        parent=styles['Normal'],
        fontName=font_name,
        fontSize=11,
        leading=14
    )

    program_html = ""
    for item in program_items:
        item = str(item).strip()
        if not item: continue
        if item.startswith("-"):
            program_html += f"<br/>• {item[1:].strip()}"
        else:
            program_html += f"<br/><br/>{item}"
    
    column_width = width/2 - 130
    frame_program = Frame(left_x, 0, column_width, top_y+45, showBoundary=0)
    p = Paragraph(program_html, style)
    frame_program.addFromList([p], c)

    # Ministrantes
    c.setFont(f"{font_name}-Bold" if font_name != "Helvetica" else "Helvetica-Bold", 11)
    c.drawCentredString(right_x, top_y-35, 'Organizadores:')
    
    c.setFont(font_name, 8)
    y2 = top_y - 16
    for m in ministrantes:
        c.drawCentredString(right_x, y2-33, str(m).strip())
        y2 -= 12
        if y2 < 60*mm: break

    # Volume e Número
    c.setFont(font_name, 12)
    if volume:
        c.drawCentredString(right_x+57, 77*mm, str(volume))
    c.drawCentredString(right_x+57, 67*mm, str(cert_num))

    # Data (se necessário na pág 2 também)
    c.drawCentredString(width / 2, 60, f"Cuiabá-MT, {date_today}")

    c.save()
    packet.seek(0)
    return packet

def generate_single_pdf(template_bytes, overlay1, overlay2):
    # Ler template da memória
    template_reader = PdfReader(template_bytes)
    writer = PdfWriter()

    # Página 1
    overlay1_pdf = PdfReader(overlay1)
    page1 = template_reader.pages[0]
    # Criar nova página em branco com dimensões do template
    new_page1 = PageObject.create_blank_page(width=page1.mediabox.width, height=page1.mediabox.height)
    new_page1.merge_page(page1)
    new_page1.merge_page(overlay1_pdf.pages[0])
    writer.add_page(new_page1)

    # Página 2 (Se existir no template)
    if len(template_reader.pages) > 1:
        overlay2_pdf = PdfReader(overlay2)
        page2 = template_reader.pages[1]
        new_page2 = PageObject.create_blank_page(width=page2.mediabox.width, height=page2.mediabox.height)
        new_page2.merge_page(page2)
        new_page2.merge_page(overlay2_pdf.pages[0])
        writer.add_page(new_page2)

    output_stream = io.BytesIO()
    writer.write(output_stream)
    return output_stream.getvalue()

# --- INTERFACE PRINCIPAL ---

st.title("🎓 Gerador de Certificados - PET Elétrica")
st.markdown("Carregue a planilha e o template PDF para gerar os certificados automaticamente.")

col1, col2 = st.columns(2)

with col1:
    uploaded_excel = st.file_uploader("📂 Carregar Planilha Excel (.xlsx)", type=["xlsx"])

with col2:
    uploaded_template = st.file_uploader("📄 Carregar Template (.pdf)", type=["pdf"])

if uploaded_excel and uploaded_template:
    try:
        xls = pd.ExcelFile(uploaded_excel)
        sheet_names = xls.sheet_names
        
        # --- SELEÇÃO DE DADOS ---
        st.divider()
        st.subheader("🛠️ Configuração do Curso")
        
        # Tenta achar a aba cursos
        curso_df = None
        if 'cursos' in sheet_names:
            curso_df = pd.read_excel(uploaded_excel, sheet_name='cursos')
            curso_df.columns = [c.upper().strip() for c in curso_df.columns]
        
        course_data = {}
        
        if curso_df is not None and not curso_df.empty:
            option_list = curso_df['CURSO'].tolist()
            selected_course_name = st.selectbox("Selecione o Curso:", option_list)
            
            # Pegar dados da linha selecionada
            row = curso_df[curso_df['CURSO'] == selected_course_name].iloc[0]
            
            # Preencher formulário com dados do Excel
            with st.expander("Ver/Editar Detalhes do Curso", expanded=True):
                c1, c2 = st.columns(2)
                course_data['NOME'] = c1.text_input("Nome do Curso", row.get('CURSO'))
                course_data['CARGA'] = c2.text_input("Carga Horária", str(row.get('CARGAHORARIA', '')))
                
                # Tratamento de datas
                d_ini = parse_date_br(row.get('DATAINICIO'))
                d_fim = parse_date_br(row.get('DATAFIM'))
                
                course_data['DT_INI'] = c1.date_input("Data Início", value=d_ini if d_ini else datetime.now())
                course_data['DT_FIM'] = c2.date_input("Data Fim", value=d_fim if d_fim else datetime.now())
                
                course_data['PROGRAMA'] = st.text_area("Conteúdo (separar por ;)", str(row.get('PROGRAMA', '')))
                course_data['MINISTRANTES'] = st.text_area("Ministrantes (separar por ;)", str(row.get('MINISTRANTES', '')))

        else:
            st.warning("Aba 'cursos' não encontrada. Preencha manualmente.")
            course_data['NOME'] = st.text_input("Nome do Curso")
            # ... (adicionar outros campos manuais se necessário)

        # --- CONFIGURAÇÃO DE GERAÇÃO ---
        st.divider()
        st.subheader("⚙️ Parâmetros da Emissão")
        
        col_p1, col_p2, col_p3 = st.columns(3)
        tipo_cert = col_p1.selectbox("Gerar para:", ["Participantes", "Ministrantes", "Ambos"])
        data_cert = col_p2.date_input("Data da Assinatura", value=datetime.now())
        num_start = col_p3.number_input("Numeração Inicial", min_value=1, value=1)
        volume_txt = col_p1.text_input("Volume (opcional)")

        # Preparar DataFrame de Pessoas
        df_final = pd.DataFrame()
        
        # Lógica de Participantes
        if tipo_cert in ["Participantes", "Ambos"]:
            sheet_part = 'participantes' if 'participantes' in sheet_names else sheet_names[0]
            df_part = pd.read_excel(uploaded_excel, sheet_name=sheet_part)
            df_part.columns = [c.upper().strip() for c in df_part.columns]
            
            # Garantir coluna TIPO
            if 'TIPO' not in df_part.columns:
                df_part['TIPO'] = 'participante'
            
            df_part['TIPO'] = df_part['TIPO'].astype(str).str.lower()
            df_final = pd.concat([df_final, df_part[df_part['TIPO'].str.contains('part')]], ignore_index=True)

        # Lógica de Ministrantes
        if tipo_cert in ["Ministrantes", "Ambos"]:
            # Ministrantes podem vir da aba cursos (string) ou da aba participantes (linhas)
            # Implementação simplificada baseada no seu código original (da string da aba cursos)
            if course_data.get('MINISTRANTES'):
                nomes_min = [m.strip() for m in course_data['MINISTRANTES'].split(';') if m.strip()]
                df_min = pd.DataFrame({'NOME': nomes_min, 'TIPO': 'ministrante'})
                df_final = pd.concat([df_final, df_min], ignore_index=True)

        st.info(f"Total de certificados a gerar: **{len(df_final)}**")

        # --- BOTÃO DE AÇÃO ---
        if st.button("🚀 Gerar Certificados", type="primary"):
            if len(df_final) == 0:
                st.error("Nenhuma pessoa encontrada para os critérios selecionados.")
            else:
                progress_bar = st.progress(0)
                zip_buffer = io.BytesIO()
                
                template_bytes = uploaded_template.getvalue()
                
                with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                    counter = num_start
                    
                    for idx, row in df_final.iterrows():
                        nome = row.get('NOME', 'Desconhecido')
                        tipo = row.get('TIPO', 'participante')
                        
                        # Definições baseadas no tipo
                        papel = 'organizou o' if 'min' in tipo else 'participou do'
                        pasta_tipo = 'Ministrantes' if 'min' in tipo else 'Participantes'
                        
                        # Texto Período
                        d_ini_str = course_data['DT_INI'].strftime('%d/%m/%Y')
                        d_fim_str = course_data['DT_FIM'].strftime('%d/%m/%Y')
                        
                        if d_ini_str == d_fim_str:
                            periodo_txt = f"realizado no dia {d_ini_str},"
                        else:
                            periodo_txt = f"realizado de {d_ini_str} a {d_fim_str},"
                            
                        # Montar HTML
                        texto_pag1 = (
                            f"Certificamos que <b>{nome}</b> {papel} <b>{course_data['NOME']}</b>, "
                            f"{periodo_txt} com carga horária de {course_data['CARGA']} horas."
                        )
                        
                        prog_list = [p.strip() for p in course_data['PROGRAMA'].split(';') if p.strip()]
                        min_list = [m.strip() for m in course_data['MINISTRANTES'].split(';') if m.strip()]
                        
                        # Gerar Overlays
                        ov1 = create_overlay_page1(texto_pag1, data_cert.strftime('%d/%m/%Y'))
                        ov2 = create_overlay_page2(prog_list, min_list, volume_txt, f"{counter:04d}", data_cert.strftime('%d/%m/%Y'))
                        
                        # Gerar PDF Final
                        pdf_bytes = generate_single_pdf(io.BytesIO(template_bytes), ov1, ov2)
                        
                        # Adicionar ao ZIP
                        nome_limpo = "".join([c for c in nome if c.isalnum() or c.isspace()]).strip()
                        filename = f"{pasta_tipo}/{counter:04d} - {nome_limpo}.pdf"
                        zip_file.writestr(filename, pdf_bytes)
                        
                        counter += 1
                        progress_bar.progress((idx + 1) / len(df_final))

                st.success("Certificados gerados com sucesso!")
                
                # Botão de Download
                st.download_button(
                    label="📥 Baixar Certificados (.zip)",
                    data=zip_buffer.getvalue(),
                    file_name=f"Certificados_{datetime.now().strftime('%Y%m%d_%H%M')}.zip",
                    mime="application/zip"
                )

    except Exception as e:
        st.error(f"Erro ao processar arquivo: {e}")
        st.write("Detalhes do erro:", e)

else:
    st.info("Aguardando upload dos arquivos...")   
