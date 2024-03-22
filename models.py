from datetime import datetime
from django.db import models
from django.shortcuts import render


class Produto(models.Model):
    nome = models.CharField(max_length=100)
    valor = models.DecimalField(max_digits=10, decimal_places=2)

    class Meta:
        db_table = 'produto'

class Precos_status_graduacao(models.Model):
    status = models.CharField(max_length=100)
    graduacao = models.CharField(max_length=100)
    valor = models.DecimalField(max_digits=10, decimal_places=2)

    class Meta:
        db_table = 'precos_por_status_e_graduacao'

class Precos_graduacao_vinculo(models.Model):
    graduacao = models.CharField(max_length=100)
    vinculo = models.CharField(max_length=100)
    valor = models.DecimalField(max_digits=10, decimal_places=2)

    class Meta:
        db_table = 'precos_por_graduacao_e_vinculo'



class BaseDados(models.Model):
    SEXO_CHOICES =[
        ('M', 'Masculino'),
        ('F', 'Feminino'),
        ('O', 'Outro'),
    ]

    STATUS_CHOICES =[
        ('CIVIL', 'CIVIL'),
        ('MILITAR DA ATIVA', 'MILITAR DA ATIVA'),
        ('MILITAR DA RESERVA', 'MILITAR DA RESERVA'),
        ('PENSIONISTA', 'PENSIONISTA'),
        ('DEP. DESACOMPANHADO', 'DEP. DESACOMPANHADO'),
    ]

    GRADUACAO_CHOICES =[
        ('GEN', 'GEN'),
        ('CEL', 'CEL'),
        ('TC', 'TC'),
        ('MAJ', 'MAJ'),
        ('CAP', 'CAP'),
        ('1º TEN', '1º TEN'),
        ('2º TEN', '2º TEN'),
        ('ASP', 'ASP'),
        ('SO', 'SO'),
        ('ST', 'ST'),
        ('1º SGT', '1º SGT'),
        ('2º SGT', '2º SGT'),
        ('3º SGT', '3º SGT'),
        ('CIVIL', 'CIVIL'),

    ]
    TIPO_CHOICES =[
        ('Casal', 'Casal'),
        ('Solteiro', 'Solteiro'),
        ('Outro', 'Outro'),
    ]
    VINCULO_CHOICES =[
        ('Cônjuge', 'Cônjuge'),
        ('Filho até 6 anos', 'Filho até 6 anos'),
        ('Filho de 7 a 10 anos', 'Filho de 7 a 10 anos'),
        ('Filho de 11 a 23 anos', 'Filho de 11 a 23 anos'),
        ('Filho acima de 23 anos', 'Filho acima de 23 anos'),
        ('Sem vínculo familiar', 'Sem vínculo familiar'),        
    ]

    QTDE_HOSP_CHOICES =[
        (1, '1'),
        (2, '2'),
        (3, '3'),
        (4, '4'),
        (5, '5'),
        (6, '6'),        
    ]

    ESPECIAL_CHOICES =[
        ('Sim', 'Sim'),
        ('Não', 'Não'),                
    ]

    STATUS_RESERVA_CHOICES =[
        ('Pendente', 'Pendente'),
        ('Aprovada', 'Aprovada'),
        ('Recusada', 'Recusada'),
        ('Concluída', 'Concluída'),                
    ]

    MOTIVO_VIAGEM_CHOICES =[
        ('Saúde', 'Saúde'),
        ('Trabalho', 'Trabalho'),
        ('Turismo', 'Turísmo'),                
    ]

    UF_CHOICES =[
        ('AL', 'AL'),
        ('AM', 'AM'),
        ('AP', 'AP'),
        ('BA', 'BA'),
        ('CE', 'CE'),
        ('DF', 'DF'),
        ('ES', 'ES'),
        ('GO', 'GO'),
        ('MA', 'MA'),
        ('MG', 'MG'),
        ('MS', 'MS'),
        ('MT', 'MT'),
        ('PA', 'PA'),
        ('PB', 'PB'),
        ('PE', 'PE'),
        ('PI', 'PI'),
        ('PR', 'PR'),
        ('RJ', 'RJ'),
        ('RN', 'RN'),
        ('RO', 'RO'),
        ('RR', 'RR'),
        ('RS', 'RS'),
        ('SC', 'SC'),
        ('SE', 'SE'),
        ('SP', 'SP'),
        ('TO', 'TO'),        
    ]
    
    id = models.AutoField(db_column='ID', primary_key=True, blank=True, null=False)
    entrada = models.DateField(db_column='ENTRADA', blank=True, null=True)
    saida = models.DateField(db_column='SAÍDA', blank=True, null=True)  
    nome = models.CharField(db_column='NOME', max_length=100, blank=True, null=True)
    diarias = models.IntegerField(db_column='DIARIAS',default=0, blank=True, null=True)   
    graduacao = models.TextField(db_column='GRADUAÇÃO',max_length=8, choices = GRADUACAO_CHOICES, blank=True, null=True)  
    telefone = models.CharField(db_column='TELEFONE', max_length=15, blank=True, null=True)  
    qtde_quartos = models.IntegerField(db_column='QTDE_QUARTOS', blank=True, null=True)  
    qtde_hosp = models.IntegerField(db_column='QTDE_HOSP',choices = QTDE_HOSP_CHOICES, blank=True, null=True)  
    especial = models.TextField(db_column='ESPECIAL', max_length=3, choices = ESPECIAL_CHOICES, blank=True, null=True)
    qtde_acomp = models.IntegerField(db_column='QTDE_ACOMP', blank=True, null=True)

    email = models.CharField(db_column='EMAIL', max_length=100, blank=True, null=True)  
    cpf = models.CharField(db_column='CPF', max_length=15, blank=True, null=True)     
    status = models.CharField(db_column='STATUS',max_length=25, choices = STATUS_CHOICES, blank=True, null=True)     
    tipo = models.TextField(db_column='TIPO',max_length=25, choices = TIPO_CHOICES, blank=True, null=True)  
    sexo = models.TextField(db_column='SEXO', max_length=1, choices = SEXO_CHOICES, blank=True, null=True)  
    cidade = models.CharField(db_column='CIDADE', max_length=100, blank=True, null=True)  
    uf = models.CharField(db_column='UF', max_length=2, choices = UF_CHOICES, blank=True, null=True)     
    status_reserva = models.CharField(db_column='STATUS_RESERVA', max_length=15, choices = STATUS_RESERVA_CHOICES, blank=True, null=True)
    nome_acomp1 = models.CharField(db_column='NOME_ACOMP1', max_length=100, blank=True, null=True)  
    vinculo_acomp1 = models.TextField(db_column='VINCULO_ACOMP1', max_length=22, choices = VINCULO_CHOICES, blank=True, null=True)  
    idade_acomp1 = models.IntegerField(db_column='IDADE_ACOMP1', blank=True, null=True)  
    sexo_acomp1 = models.TextField(db_column='SEXO_ACOMP1', max_length=1, choices = SEXO_CHOICES, blank=True, null=True)  
    nome_acomp2 = models.CharField(db_column='NOME_ACOMP2', max_length=100, blank=True, null=True)  
    vinculo_acomp2 = models.TextField(db_column='VINCULO_ACOMP2', max_length=22, choices = VINCULO_CHOICES, blank=True, null=True)  
    idade_acomp2 = models.IntegerField(db_column='IDADE_ACOMP2', blank=True, null=True)  
    sexo_acomp2 = models.TextField(db_column='SEXO_ACOMP2', max_length=1, choices = SEXO_CHOICES, blank=True, null=True)  
    nome_acomp3 = models.CharField(db_column='NOME_ACOMP3', max_length=100, blank=True, null=True)  
    vinculo_acomp3 = models.TextField(db_column='VINCULO_ACOMP3', max_length=22, choices = VINCULO_CHOICES, blank=True, null=True)  
    idade_acomp3 = models.IntegerField(db_column='IDADE_ACOMP3', blank=True, null=True)  
    sexo_acomp3 = models.TextField(db_column='SEXO_ACOMP3', max_length=1, choices = SEXO_CHOICES, blank=True, null=True)  
    nome_acomp4 = models.CharField(db_column='NOME_ACOMP4', max_length=100, blank=True, null=True)  
    vinculo_acomp4 = models.TextField(db_column='VINCULO_ACOMP4', max_length=22, choices = VINCULO_CHOICES, blank=True, null=True)  
    idade_acomp4 = models.IntegerField(db_column='IDADE_ACOMP4', blank=True, null=True)  
    sexo_acomp4 = models.TextField(db_column='SEXO_ACOMP4', max_length=1, choices = SEXO_CHOICES, blank=True, null=True)  
    nome_acomp5 = models.CharField(db_column='NOME_ACOMP5', max_length=100, blank=True, null=True)  
    vinculo_acomp5 = models.TextField(db_column='VINCULO_ACOMP5', max_length=22, choices = VINCULO_CHOICES, blank=True, null=True)  
    idade_acomp5 = models.IntegerField(db_column='IDADE_ACOMP5', blank=True, null=True)  
    sexo_acomp5 = models.TextField(db_column='SEXO_ACOMP5', max_length=1, choices = SEXO_CHOICES, blank=True, null=True)  

    mhex = models.CharField(db_column='MHEx', max_length=6, blank=True, null=True)  
    uh = models.CharField(db_column='UH', max_length=2, blank=True, null=True) 
    forma_pagamento = models.CharField(db_column='FORMA_PAGAMENTO', max_length=2, blank=True, null=True)

    valor_hosp = models.DecimalField(db_column='VALOR_HOSP', max_digits=10, decimal_places=2, default=0)
    valor_acomp1 = models.DecimalField(db_column='VALOR_ACOMP1', max_digits=10, decimal_places=2, default=0)
    valor_acomp2 = models.DecimalField(db_column='VALOR_ACOMP2', max_digits=10, decimal_places=2, default=0)
    valor_acomp3 = models.DecimalField(db_column='VALOR_ACOMP3', max_digits=10, decimal_places=2, default=0)
    valor_acomp4 = models.DecimalField(db_column='VALOR_ACOMP4', max_digits=10, decimal_places=2, default=0)
    valor_acomp5 = models.DecimalField(db_column='VALOR_ACOMP5', max_digits=10, decimal_places=2, default=0)

    valor_dia = models.DecimalField(db_column='VALOR_DIA', max_digits=10, decimal_places=2, default=0)
    valor_ajuste = models.DecimalField(db_column='VALOR_AJUSTE', max_digits=10, decimal_places=2, default=0)  
    subtotal = models.DecimalField(db_column='SUBTOTAL', max_digits=10, decimal_places=2, default=0) 
    valor_total = models.DecimalField(db_column='VALOR_TOTAL', max_digits=10, decimal_places=2, default=0) 

    qtde_agua = models.IntegerField(db_column='QTDE_AGUA', blank=True, null=True)
    qtde_refri = models.IntegerField(db_column='QTDE_REFRI', blank=True, null=True)  
    qtde_cerveja = models.IntegerField(db_column='QTDE_CERVEJA', blank=True, null=True)
    total_agua = models.DecimalField(db_column='TOTAL_AGUA', max_digits=10, decimal_places=2, default=0) 
    total_refri = models.DecimalField(db_column='TOTAL_REFRI', max_digits=10, decimal_places=2, default=0)
    total_cerveja = models.DecimalField(db_column='TOTAL_CERVEJA', max_digits=10, decimal_places=2, default=0)
    total_consumacao = models.DecimalField(db_column='TOTAL_CONSUMACAO', max_digits=10, decimal_places=2, default=0)       
   
    nome_pagante = models.CharField(db_column='NOME_PAGANTE', max_length=100, blank=True, null=True)
    cpf_pagante = models.CharField(db_column='CPF_PAGANTE', max_length=100, blank=True, null=True)

    motivo_viagem = models.TextField(db_column='MOTIVO_VIAGEM',choices = MOTIVO_VIAGEM_CHOICES, max_length=100, blank=True, null=True)
    desc_saude = models.DecimalField(db_column='DESC_SAUDE', max_digits=10, decimal_places=2, default=0)

    #diáriastotal = models.TextField(db_column='DIÁRIASTOTAL', blank=True, null=True)  
    #água = models.TextField(db_column='ÁGUA', blank=True, null=True)  
    #refri = models.TextField(db_column='REFRI', blank=True, null=True)  
    #cerveja = models.TextField(db_column='CERVEJA', blank=True, null=True)  
    #consumação = models.TextField(db_column='CONSUMAÇÃO', blank=True, null=True)  
    #ajustes = models.TextField(db_column='AJUSTES', blank=True, null=True)  
    #totalgeral = models.TextField(db_column='TOTALGERAL', blank=True, null=True)  
    #média = models.TextField(db_column='Média', blank=True, null=True)  
    #pagamento = models.TextField(db_column='PAGAMENTO', blank=True, null=True)  
    
        
    
    class Meta:
        db_table = 'base_dados'

