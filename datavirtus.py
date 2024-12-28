import docx
import faker

class RelatorioVirtus():

    def __init__(self, arquivo):
        self.arquivo = arquivo
        self.texto = self.ler_arquivo(arquivo)

    def ler_arquivo(self, arquivo):
        doc = docx.Document(arquivo)
        texto = []
        for paragrafo in doc.paragraphs:
            texto.append(paragrafo.text)
        return texto

def cpf_or_cnpj(cpf_cnpj):
    cpf_cnpj = ''.join(filter(str.isdigit, str(cpf_cnpj)))
    if len(cpf_cnpj) == 11:
        return 'CPF'
    elif len(cpf_cnpj) == 14:
        return 'CNPJ'
    else:
        return 'Inválido'

def anonimizar(dataframe, coluna, tipo, seed=None):

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