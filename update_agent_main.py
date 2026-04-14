import sys

content = open('src/local_agent/main.py', 'r', encoding='utf-8').read()

import_code = '''try:
    from src.tabs.tab_automated_actions import (
        worker_mover_por_coincidencia,
        worker_organizar_facturas_feov,
        worker_copiar_mapeo_subcarpetas,
        worker_copiar_archivos_desde_raiz_mapeo
    )
except ImportError:
    pass
'''

if 'worker_mover_por_coincidencia' not in content:
    content = content.replace('import traceback', 'import traceback\n' + import_code)

agent_tasks = '''
            elif command == "mover_por_coincidencia":
                root_path = params.get("root_path")
                if root_path:
                    self.log(f"Mover por coincidencia en: {root_path}")
                    res = worker_mover_por_coincidencia(root_path, silent_mode=True, return_zip=False)
                    result["result"] = res
                else:
                    result["status"] = "ERROR"
                    result["result"] = {"error": "Falta root_path"}

            elif command == "organizar_feov":
                root_path = params.get("root_path")
                target_path = params.get("target_path")
                if root_path and target_path:
                    self.log(f"Organizar FEOV: {root_path} -> {target_path}")
                    res = worker_organizar_facturas_feov(root_path, target_path, silent_mode=True, return_zip=False)
                    result["result"] = res
                else:
                    result["status"] = "ERROR"
                    result["result"] = {"error": "Faltan rutas"}

            elif command == "copiar_mapeo_subcarpetas":
                b64_file = params.get("file_bytes_b64")
                sheet_name = params.get("sheet_name")
                col_src = params.get("col_src")
                col_dst = params.get("col_dst")
                path_src_base = params.get("path_src_base")
                path_dst_base = params.get("path_dst_base")
                use_filter = params.get("use_filter", False)
                if b64_file and path_src_base and path_dst_base:
                    file_bytes = base64.b64decode(b64_file)
                    self.log(f"Copiar mapeo subcarpetas a: {path_dst_base}")
                    res = worker_copiar_mapeo_subcarpetas(file_bytes, sheet_name, col_src, col_dst, path_src_base, path_dst_base, use_filter, silent_mode=True)
                    result["result"] = {"message": res}
                else:
                    result["status"] = "ERROR"
                    result["result"] = {"error": "Faltan parámetros"}

            elif command == "copiar_archivos_desde_raiz_mapeo":
                b64_file = params.get("file_bytes_b64")
                sheet_name = params.get("sheet_name")
                col_id = params.get("col_id")
                col_folder = params.get("col_folder")
                root_src = params.get("root_src")
                root_dst = params.get("root_dst")
                use_filter = params.get("use_filter", False)
                if b64_file and root_src and root_dst:
                    file_bytes = base64.b64decode(b64_file)
                    self.log(f"Copiar desde raiz mapeo a: {root_dst}")
                    res = worker_copiar_archivos_desde_raiz_mapeo(file_bytes, sheet_name, col_id, col_folder, root_src, root_dst, use_filter, silent_mode=True)
                    result["result"] = {"message": res}
                else:
                    result["status"] = "ERROR"
                    result["result"] = {"error": "Faltan parámetros"}
'''

if 'elif command == "mover_por_coincidencia":' not in content:
    content = content.replace('            elif command == "flat_to_excel":', agent_tasks + '\n            elif command == "flat_to_excel":')

with open('src/local_agent/main.py', 'w', encoding='utf-8') as f:
    f.write(content)
