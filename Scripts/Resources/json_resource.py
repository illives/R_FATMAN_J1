
import json

class json_tool:

    def __init__(self, data_input, data_output, json_path):
        self.data_input = data_input
        self.data_output = data_output
        self.json_path = json_path

    def save_json(data_input, json_path):
        """
        Transforma dados em arquivo JSON.
        Todos os argumentos são obrigatorios.
        Args:
        (data_input = em formato de dict{},
        jason_path = caminho rede onde ficar o arquivo Json.)
        """   
        with open(json_path, "w") as f:
            json.dump(data_input, f, ensure_ascii=False, indent=4)

    def read_json(json_path):
        """
        Lê os dados de um arquivo JSON.
        Todos os argumentos são obrigatorios.
        Args:
        (jason_path = caminho rede onde ficar o arquivo Json.)
        """  
        with open(json_path) as f:
            data = json.load(f)
            return data