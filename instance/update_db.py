from app import app, db

# Renomear tabela existente para um nome temporário
db.engine.execute('ALTER TABLE ticket RENAME TO ticket_old;')

# Criar uma nova tabela com a estrutura desejada
db.engine.execute('''
CREATE TABLE ticket (
    id STRING PRIMARY KEY,
    username STRING NOT NULL,
    tool STRING NOT NULL,
    subject STRING NOT NULL,
    description TEXT NOT NULL,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
);
''')

# Copiar dados da tabela antiga para a nova tabela
db.engine.execute('''
INSERT INTO ticket (id, username, tool, subject, description, created_at)
SELECT id, username, tool, subject, description, created_at FROM ticket_old;
''')

# Excluir a tabela antiga
db.engine.execute('DROP TABLE ticket_old;')

print("Database updated successfully.")