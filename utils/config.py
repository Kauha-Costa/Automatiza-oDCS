import json
import os

CONFIG_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "config.json")

DEFAULT_CONFIG = {
    "pasta_certificados": "",
    "imprimir_apos": False,
    "impressora": ""
}


def carregar_config() -> dict:
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                dados = json.load(f)
                # Garante que todas as chaves existam mesmo em configs antigas
                for k, v in DEFAULT_CONFIG.items():
                    dados.setdefault(k, v)
                return dados
        except Exception:
            pass
    return DEFAULT_CONFIG.copy()


def salvar_config(config: dict):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"Aviso: não foi possível salvar configuração: {e}")
