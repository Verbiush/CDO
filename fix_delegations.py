import re

with open('src/tabs/tab_automated_actions.py', 'r', encoding='utf-8') as f:
    content = f.read()

pattern = re.compile(
    r'res = wait_for_result\(task_id, timeout=(\d+)\)\s+'
    r'if res and res\.get\(\"status\"\) == \"COMPLETED\":\s+'
    r'return res\.get\(\"result\"\)\s+'
    r'else:\s+'
    r'return \{\"error\": f\"Error en agente: \{res\.get\(\'error\'\) if res else \'Sin respuesta\'\}\"\}'
)

def replace_func(m):
    timeout = m.group(1)
    return (
        f'res = wait_for_result(task_id, timeout={timeout})\n'
        f'            if res and "error" not in res:\n'
        f'                return res\n'
        f'            else:\n'
        f'                return {{"error": f"Error en agente: {{res.get(\'error\') if res else \'Sin respuesta\'}}"}}'
    )

new_content = pattern.sub(replace_func, content)

with open('src/tabs/tab_automated_actions.py', 'w', encoding='utf-8') as f:
    f.write(new_content)

print('Replaced instances:', len(pattern.findall(content)))
