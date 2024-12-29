import docx
import faker
import re
import json
from docx.shared import Inches


def cpf_or_cnpj(cpf_cnpj):
    cpf_cnpj = ''.join(filter(str.isdigit, str(cpf_cnpj)))
    if len(cpf_cnpj) == 11:
        return 'CPF'
    elif len(cpf_cnpj) == 14:
        return 'CNPJ'
    else:
        return 'Inválido'


def pseudoanonimizar(dataframe, coluna, tipo, seed=None):

    if seed is not None:
        faker.Faker.seed(seed)

    fake = faker.Faker('pt_BR')

    reais_unicos = dataframe[coluna].unique()
    if tipo == 'nome':
        fake_values = list(map(lambda _: fake.name(), range(len(reais_unicos))))
    elif tipo == 'email':
        fake_values = list(map(lambda _: fake.email(), range(len(reais_unicos))))
    elif tipo == 'cpf_cnpj':
        fake_values = []
        for valor in reais_unicos:
            tipo_documento = cpf_or_cnpj(valor)
            if tipo_documento == 'CPF':
                fake_values.append(fake.cpf())
            elif tipo_documento == 'CNPJ':
                fake_values.append(fake.cnpj())
    elif tipo == 'telefone':
        fake_values = list(map(lambda _: fake.phone_number(), range(len(reais_unicos))))
    elif tipo == 'endereco':
        fake_values = list(map(lambda _: fake.address(), range(len(reais_unicos))))
    elif tipo == 'data':
        fake_values = list(map(lambda _: fake.date(), range(len(reais_unicos))))
    else:
        raise ValueError('Tipo inválido')

    fake_dict = dict(zip(reais_unicos, fake_values))
    dataframe[coluna] = dataframe[coluna].replace(fake_dict)

    return dataframe


class RelatorioVirtus:

    def __init__(self, modelo_relatorio, lista_tags=None):
        self.modelo_relatorio = modelo_relatorio
        self.conteudo = self.ler_arquivo()
        if lista_tags is not None:
            self.mapa_tags = self.carregar_tags(lista_tags)
        else:
            self.extrair_tags()  # Extrai as tags do texto do arquivo docx

    def ler_arquivo(self):
        doc = docx.Document(self.modelo_relatorio)
        conteudo = []
        for paragrafo in doc.paragraphs:
            conteudo.append(paragrafo.text)
        return conteudo

    def extrair_tags(self):
        tags = set()
        for paragraph in self.conteudo:
            tags.update(re.findall(r'(\|.*?\|)', paragraph))
        self.mapa_tags = {tag: None for tag in tags}
        self.ordenar_tags()

    def ordenar_tags(self):
        self.mapa_tags = dict(sorted(self.mapa_tags.items(), key=lambda x: x[0]))

    def exportar_tags(self, arquivo_saida=None):

        if arquivo_saida is None:
            arquivo_saida = 'mapa_tags.json'

        with open(arquivo_saida, 'w', encoding='utf-8') as f:
            json.dump(self.mapa_tags, f, ensure_ascii=False, indent=4)

    def carregar_tags(self, arquivo_tags='mapa_tags.json'):
        with open(arquivo_tags, 'r', encoding='utf-8') as f:
            self.mapa_tags = json.load(f)
        self.ordenar_tags()
        return self.mapa_tags

    def substituir_tags(self):
        for i, paragraph in enumerate(self.conteudo):
            for tag, valor in self.mapa_tags.items():
                if valor is not None:
                    self.conteudo[i] = self.conteudo[i].replace(tag, valor)
        return self.conteudo

    def salvar_relatorio(self, arquivo_saida='relatorio.docx'):

        doc = docx.Document()
        for paragraph in self.conteudo:
            doc.add_paragraph(paragraph)
        doc.save(arquivo_saida)
        return arquivo_saida

    def substituir_grafico(self, imagem_grafico, tag_grafico):
        doc = docx.Document()
        for paragraph in self.conteudo:
            if tag_grafico in paragraph:
                paragraph = paragraph.replace(tag_grafico, '')
                doc.add_paragraph(paragraph)
                doc.add_picture(imagem_grafico, width=Inches(6.0))  # Adjust the width as needed
            else:
                doc.add_paragraph(paragraph)
        self.conteudo = [p.text for p in doc.paragraphs]
        doc.save('relatorio_final.docx')
        return self.conteudo
