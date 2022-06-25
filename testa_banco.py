import mysql.connector

conexao = mysql.connector.connect(
    host = "localhost",
    user = "root",
    passwd = "",
    database = "transportadora"
)

cursor = conexao.cursor()
cursor.execute('''CREATE TABLE IF NOT EXISTS clientes(
                id_cliente INT NOT NULL AUTO_INCREMENT,
                cpf CHAR(14) NOT NULL, 
                nome VARCHAR(100) NOT NULL,
                senha TEXT(30) NOT NULL,
                numero_do_cartao VARCHAR(16) NOT NULL,
                email TEXT(256) NOT NULL,
                ddd INT(2) NOT NULL,
                numero BIGINT(9) NOT NULL,
                PRIMARY KEY(id_cliente))default charset = utf8;''')
cursor.execute('''CREATE TABLE IF NOT EXISTS mercadorias(
                id_mercadoria INT NOT NULL AUTO_INCREMENT,
                id_fk_cliente INT,
                tipo_de_carga CHAR NOT NULL, 
                comprimento INT(3) NOT NULL,
                altura INT(3) NOT NULL,
                largura INT(3) NOT NULL,
                quantidade INT(4) NOT NULL,
                PRIMARY KEY(id_mercadoria),
                FOREIGN KEY(id_fk_cliente) REFERENCES clientes(id_cliente))default charset = utf8;''')
cursor.execute('''CREATE TABLE IF NOT EXISTS localizacao(
                id_localizacao INT NOT NULL AUTO_INCREMENT,
                fk_id_cliente INT,
                cep VARCHAR(8) NOT NULL,
                bairro VARCHAR(50) NOT NULL,
                rua VARCHAR(255) NOT NULL,
                quadra INT(3),
                lote INT(3),
                PRIMARY KEY(id_localizacao),
                FOREIGN KEY(fk_id_cliente) REFERENCES clientes(id_cliente))default charset = utf8;''')
cursor.execute('''CREATE TABLE IF NOT EXISTS boletos(
                id_boleto INT NOT NULL AUTO_INCREMENT,
                id_fk_cliente INT,
                valor FLOAT,
                data DATE,
                fk_nome_cliente VARCHAR(100),
                fk_quantidade INT(4),
                fk_tipo_carga CHAR,
                fk_cep VARCHAR(8),
                fk_bairro VARCHAR(50),
                fk_rua VARCHAR(255),
                fk_quadra INT(3),
                fk_lote INT(3),
                PRIMARY KEY(id_boleto),
                FOREIGN KEY(fk_quantidade) REFERENCES mercadorias(quantidade),
                FOREIGN KEY(fk_tipo_carga) REFERENCES mercadorias(tipo_de_carga),
                FOREIGN KEY(id_fk_cliente) REFERENCES clientes(id_cliente),
                FOREIGN KEY(fk_cep) REFERENCES localizacao(cep),
                FOREIGN KEY(fk_bairro) REFERENCES mercadorias(bairro),
                FOREIGN KEY(fk_rua) REFERENCES mercadorias(rua),
                FOREIGN KEY(fk_quadra) REFERENCES mercadorias(quadra),
                FOREIGN KEY(fk_lote) REFERENCES mercadorias(lote),
                FOREIGN KEY(fk_nome_cliente) REFERENCES clientes(nome))default charset = utf8;''')
cursor.execute('''CREATE TABLE IF NOT EXISTS caminhao(
                id_caminhao INT NOT NULL AUTO_INCREMENT,
                fk_cep_cliente VARCHAR(8),
                fk_mercadoria INT,
                cep_caminhao VARCHAR(8) NOT NULL,
                PRIMARY KEY(id_caminhao),
                FOREIGN KEY(fk_cep_cliente) REFERENCES localizacao(cep),
                FOREIGN KEY(fk_mercadoria) REFERENCES mercadorias(id_mercadoria))default charset = utf8;''')
print(conexao)