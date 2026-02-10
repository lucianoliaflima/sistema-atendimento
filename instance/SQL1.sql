Execução finalizada com erros.
Resultado: no such table: tickets
Na linha 1:
UPDATE tickets
SET created_at = strftime('%Y-%m-%d %H:%M:%S', 'now')
WHERE created_at > strftime('%Y-%m-%d %H:%M:%S', 'now');