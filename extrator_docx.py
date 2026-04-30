"""
Extrai dados estruturados de um contrato social (DOCX ou PDF).
Se OPENAI_API_KEY estiver disponível, usa GPT-4o mini para extração completa.
Sem GPT: extração básica por regex.
"""

import json
import os
import re
from docx import Document


# ---------------------------------------------------------------------------
# Extração de texto
# ---------------------------------------------------------------------------

def extrair_texto_docx(arquivo) -> str:
    doc = Document(arquivo)
    linhas = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    return "\n".join(linhas)


def extrair_texto_pdf(arquivo) -> str:
    import pdfplumber
    texto = []
    with pdfplumber.open(arquivo) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                texto.append(t)
    return "\n".join(texto)


def extrair_texto(arquivo, nome_arquivo: str) -> str:
    ext = os.path.splitext(nome_arquivo)[1].lower()
    if ext == ".pdf":
        return extrair_texto_pdf(arquivo)
    return extrair_texto_docx(arquivo)


# ---------------------------------------------------------------------------
# Extração via GPT-4o mini
# ---------------------------------------------------------------------------

def extrair_com_gpt(texto: str) -> dict:
    api_key = os.environ.get("OPENAI_API_KEY", "")
    if not api_key:
        return {}

    import urllib.request

    prompt = """Leia o contrato social ou alteração contratual abaixo e extraia os dados no formato JSON exato:
{
  "razaoSocial": "",
  "cnpj": "",
  "nire": "",
  "objetoSocial": "",
  "capitalSocial": 0,
  "classificacao": "me|epp|regime_normal",
  "dataInicio": "",
  "numero_alteracao": 0,
  "enderecoComercial": {
    "logradouroTipo": "", "logradouroDescricao": "", "numero": "",
    "complemento": "", "bairro": "", "cidade": "", "estado": "", "cep": ""
  },
  "atividades": [
    {"cnae": "", "descricao": "", "principal": true, "desenvolvidaNoLocal": true}
  ],
  "socios": [
    {
      "nome": "", "cpf": "", "dataNascimento": "", "genero": "masculino|feminino",
      "nacionalidade": "brasileira", "estadoCivil": "solteiro|casado|divorciado|viuvo",
      "regimeBens": "", "profissao": "",
      "endereco": {
        "logradouroTipo": "", "logradouroDescricao": "", "numero": "",
        "complemento": "", "bairro": "", "cidade": "", "estado": "", "cep": ""
      },
      "documentoIdentificacao": {
        "tipo": "rg|cnh", "numero": "", "orgaoExpedidor": "", "dataExpedicao": ""
      },
      "quantidadeCotas": 0, "valorUnitarioCota": 1,
      "administrador": false, "tipoAdministracao": "isolada"
    }
  ]
}
Regras:
- Retorne APENAS o JSON válido, sem explicações nem markdown.
- Campos não encontrados: string vazia ou 0.
- dataNascimento e dataExpedicao: formato YYYY-MM-DD.
- dataInicio (início das atividades): formato DD/MM/YYYY.
- classificacao: "me" se mencionar Microempresa/ME, "epp" se EPP, senão "regime_normal".
- Se for alteração contratual consolidada, use os dados do contrato consolidado ao final.
- Para o endereço: separe corretamente logradouroTipo (Rua, Avenida...) do logradouroDescricao.
- numero_alteracao: se o documento for uma alteração contratual, retorne o número ordinal dela (ex: "SEGUNDA ALTERAÇÃO" → 2, "SEXTA ALTERAÇÃO" → 6). Se for contrato social original (constituição), retorne 0.
- nire: SEMPRE procure o NIRE em DOIS locais possíveis: (a) no cabeçalho de uma alteração contratual anterior, junto com CNPJ/razão social; (b) na parte inferior/rodapé/final de um contrato social, geralmente próximo ao registro da Junta Comercial. Aceite formatos com pontos, traços ou espaços (ex: "41.2.0123456-7", "412.0123456-7", "41 2 0123456 7"). Retorne apenas os dígitos.
- regimeBens: OBRIGATÓRIO quando estadoCivil for "casado". Procure no documento o regime de casamento de cada sócio casado e mapeie para um destes códigos: "comunhao_parcial" (comunhão parcial de bens), "comunhao_universal" (comunhão universal de bens), "separacao_total" (separação total/absoluta de bens), "separacao_obrigatoria" (separação obrigatória/legal), "participacao_final_aquestos" (participação final nos aquestos). Nunca deixe em branco se o sócio for casado.
"""

    try:
        payload = json.dumps({
            "model": "gpt-4o-mini",
            "messages": [
                {"role": "system", "content": prompt},
                {"role": "user", "content": texto[:10000]}
            ],
            "max_tokens": 2500,
            "temperature": 0.1,
            "response_format": {"type": "json_object"}
        }).encode()

        req = urllib.request.Request(
            "https://api.openai.com/v1/chat/completions",
            data=payload,
            headers={
                "Content-Type": "application/json",
                "Authorization": f"Bearer {api_key}"
            }
        )
        with urllib.request.urlopen(req, timeout=45) as resp:
            resultado = json.loads(resp.read())
            return json.loads(resultado["choices"][0]["message"]["content"])
    except Exception as e:
        print(f"[extrator] GPT error: {e}")
        return {}


# ---------------------------------------------------------------------------
# Extração local por regex (fallback sem GPT)
# ---------------------------------------------------------------------------

def extrair_local(texto: str) -> dict:
    dados = {
        "razaoSocial": "", "cnpj": "", "nire": "",
        "objetoSocial": "", "capitalSocial": 0, "classificacao": "me",
        "dataInicio": "", "numero_alteracao": 0,
        "enderecoComercial": {
            "logradouroTipo": "Rua", "logradouroDescricao": "", "numero": "",
            "complemento": "", "bairro": "", "cidade": "", "estado": "PR", "cep": ""
        },
        "atividades": [], "socios": []
    }

    # Número da alteração — detecta "SEGUNDA ALTERAÇÃO", "3ª ALTERAÇÃO", "TERCEIRA ALTERAÇÃO", etc.
    _ord_map = {
        "primeira": 1, "segunda": 2, "terceira": 3, "quarta": 4,
        "quinta": 5, "sexta": 6, "sétima": 7, "setima": 7,
        "oitava": 8, "nona": 9, "décima": 10, "decima": 10,
    }
    m_alt = re.search(r'(\d+)ª?\s*alteração\s+(?:do\s+)?contrato', texto, re.IGNORECASE)
    if m_alt:
        dados["numero_alteracao"] = int(m_alt.group(1))
    else:
        m_alt = re.search(r'(primeira|segunda|terceira|quarta|quinta|sexta|sétima|setima|oitava|nona|décima|decima)\s+alteração\s+(?:do\s+)?contrato', texto, re.IGNORECASE)
        if m_alt:
            dados["numero_alteracao"] = _ord_map.get(m_alt.group(1).lower(), 0)

    # CNPJ
    m = re.search(r'CNPJ[:\s]+([\d\.\-\/]+)', texto, re.IGNORECASE)
    if m:
        dados["cnpj"] = m.group(1).strip()

    # NIRE — pode aparecer no cabeçalho de alteração ou no rodapé do contrato social,
    # com formatos variados (pontos, traços, espaços)
    m = re.search(r'NIRE[:\s\-]*([\d][\d\.\s\-]{5,})', texto, re.IGNORECASE)
    if m:
        dados["nire"] = re.sub(r'\D', '', m.group(1))

    # Razão social — "gira sob o nome de X" ou "nome empresarial: X" ou "sob o nome de X"
    for padrao in [
        r'(?:gira|firma)\s+sob\s+o\s+nome\s+de\s+([\w\s\-\.]+(?:LTDA|EIRELI|S\.A\.|ME|EPP|UNIPESSOAL))',
        r'nome\s+empresarial[:\s]+([\w\s\-\.]+(?:LTDA|EIRELI|S\.A\.|ME|EPP|UNIPESSOAL))',
        r'sociedade[^\n]+nome\s+de\s+([\w\s\-\.]+(?:LTDA|EIRELI|S\.A\.|ME|EPP))',
    ]:
        m = re.search(padrao, texto, re.IGNORECASE)
        if m:
            dados["razaoSocial"] = m.group(1).strip().rstrip(".,;")
            break

    # Capital social
    for padrao in [
        r'capital(?:\s+social)?\s+(?:é\s+de|de|será\s+de)\s+R\$\s*([\d\.,]+)',
        r'R\$\s*([\d\.,]+)\s*\([^)]+\)[^,\n]+capital',
    ]:
        m = re.search(padrao, texto, re.IGNORECASE)
        if m:
            val = m.group(1).replace(".", "").replace(",", ".")
            try:
                dados["capitalSocial"] = float(val)
                break
            except Exception:
                pass

    # Sede / cidade
    m = re.search(r'sede[^\n]+(?:em|na|no)\s+([A-ZÀ-Ú][a-zà-ú]+(?:\s+[A-ZÀ-Ú][a-zà-ú]+)*)\s*[-–]\s*([A-Z]{2})', texto)
    if m:
        dados["enderecoComercial"]["cidade"] = m.group(1)
        dados["enderecoComercial"]["estado"] = m.group(2)

    # CEP
    m = re.search(r'CEP[:\s]*([\d]{5}[-\.]?[\d]{3})', texto, re.IGNORECASE)
    if m:
        dados["enderecoComercial"]["cep"] = m.group(1).replace(".", "").replace("-", "")

    # Classificação
    if re.search(r'microempresa|\bME\b', texto, re.IGNORECASE):
        dados["classificacao"] = "me"
    elif re.search(r'pequeno porte|\bEPP\b', texto, re.IGNORECASE):
        dados["classificacao"] = "epp"
    else:
        dados["classificacao"] = "regime_normal"

    # CNAEs
    for m in re.finditer(r'CNAE[^\d]*([\d]{4}[-\./][\d]{1,2}/[\d]{2})[^\-\n]*[-–]\s*([^\n]+)', texto, re.IGNORECASE):
        dados["atividades"].append({
            "cnae": m.group(1).strip(),
            "descricao": m.group(2).strip().rstrip(".,;"),
            "principal": len(dados["atividades"]) == 0,
            "desenvolvidaNoLocal": True,
        })

    return dados


# ---------------------------------------------------------------------------
# Função principal
# ---------------------------------------------------------------------------

def extrair_dados_contrato(arquivo, nome_arquivo: str = "") -> dict:
    """
    Extrai dados de um contrato social em DOCX ou PDF.
    - Com OPENAI_API_KEY: usa GPT-4o mini (resultado completo).
    - Sem chave: regex básico (dados essenciais).
    """
    if not nome_arquivo and hasattr(arquivo, "filename"):
        nome_arquivo = arquivo.filename or ""

    texto = extrair_texto(arquivo, nome_arquivo)

    api_key = os.environ.get("OPENAI_API_KEY", "")
    if api_key:
        dados = extrair_com_gpt(texto)
        if dados:
            return dados

    return extrair_local(texto)
