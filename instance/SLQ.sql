num_tickets = 100

sql_script = "INSERT INTO ticket (username, tool, subject, description, created_at, status, closed_at, CSR) VALUES\n"

for i in range(1, num_tickets + 1):
    username = f"user{i}"
    tool = f"Ferramenta{i % 10 + 1}"
    subject = f"Assunto{i}"
    description = f"Descrição{i}"
    day = (i % 28) + 1  # Para garantir que o dia esteja no intervalo 1-28
    created_at = f"2024-01-{day:02d} 12:00:00"
    status = "open" if i % 2 == 0 else "closed"
    closed_day = (day + 1) if (day + 1) <= 28 else 1  # Para garantir que a data de fechamento também seja válida
    closed_at = "NULL" if status == "open" else f"'2024-01-{closed_day:02d} 12:00:00'"
    csr = f"CSR{i % 10 + 1}"

    sql_script += f"('{username}', '{tool}', '{subject}', '{description}', '{created_at}', '{status}', {closed_at}, '{csr}')"
    
    if i < num_tickets:
        sql_script += ",\n"
    else:
        sql_script += ";\n"

# Caminho onde o arquivo SQL será salvo
sql_file_path = "D:\\Cod_Python\\Abertura_de_Atendimento\\instance\\SQL.sql"

# Salva o script em um arquivo
with open(sql_file_path, "w") as file:
    file.write(sql_script)

print(f"Script SQL gerado com sucesso e salvo em: {sql_file_path}")
