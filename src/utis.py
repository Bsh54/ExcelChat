from functools import wraps
import pickle
import ast
import re
import sys
import os


def withoutconnect(func):
    @wraps(func)
    def wrapper(self, *args, **kwargs):
        try:
            self.itemChanged.disconnect()
        except:
            pass

        # On essaie d'appeler la fonction normalement.
        # Si c'est un signal Qt qui envoie un argument de trop (comme un booléen), on réessaie sans les args.
        try:
            r = func(self, *args, **kwargs)
        except TypeError:
            r = func(self)

        try:
            self.itemChanged.connect(self.handle_item_changed)
        except:
            pass
        return r

    return wrapper


def translate_to_conversational(full_response):
    """Extrait l'explication conversationnelle de la réponse de l'IA."""
    try:
        # On sépare la conversation du code
        if "[CODE]" in full_response:
            parts = full_response.split("[CODE]")
            explanation = parts[0].strip()
            code_part = parts[1]
        else:
            # Si l'IA a oublié la balise mais a écrit du code
            if "def process_data" in full_response:
                explanation = full_response.split("def process_data")[0].strip()
                code_part = full_response[full_response.find("def process_data"):]
            else:
                return full_response, ""

        # Heuristiques pour la formule
        formula = ""
        if ".mean()" in code_part: formula = "=MOYENNE(...)"
        elif ".sum()" in code_part: formula = "=SOMME(...)"
        elif ".sort_values" in code_part: formula = "Tri personnalisé"
        elif "==" in code_part or ">" in code_part or "<" in code_part: formula = "=FILTRE(...)"
        elif "np.random" in code_part: formula = "=ALEA()"

        # Si l'explication est vide, on en génère une par défaut
        if not explanation:
            explanation = "Je traite votre demande sur le tableau..."

        return explanation, formula
    except:
        return "Analyse en cours...", ""

def wrap_code(full_response, df):
    # On extrait uniquement la partie technique entre [CODE] et [/CODE]
    if "[CODE]" in full_response:
        code = full_response.split("[CODE]")[1]
        # On supprime la balise de fin si elle existe, ainsi que tout ce qui suit
        if "[/CODE]" in code:
            code = code.split("[/CODE]")[0]
        code = code.strip()
    elif "def process_data" in full_response:
        code = full_response[full_response.find("def process_data"):]
    else:
        code = full_response

    # Nettoyage profond pour enlever les résidus markdown et techniques
    if "#A:" in code:
        code = code.split("#A:")[1]

    # Nettoyage ligne par ligne
    lines = code.split('\n')
    clean_lines = []
    start_collecting = False

    for line in lines:
        # On ignore les balises markdown et les lignes de langage
        if "```" in line or line.strip().lower() == "python" or "[/CODE]" in line:
            continue

        if re.match(r"^\s*(def|import|from|df|#|_data|process_data)", line):
            start_collecting = True

        if start_collecting:
            clean_lines.append(line)

    code = "\n".join(clean_lines).strip()
    if not code: return ""

    # S'assurer que la fonction est bien définie
    if re.match(r"^[a-zA-Z_]\w*\(df\):", code):
        code = "def " + code

    pattern = r"def\s+([a-zA-Z_][a-zA-Z0-9_]*)\s*\("
    match = re.search(pattern, code)

    if not match:
        function_name = "process_data"
        indented_code = "\n".join(["    " + line for line in code.split("\n") if line.strip()])
        code = f"def {function_name}(df):\n{indented_code}"
    else:
        function_name = match.group(1)

    import_line = 'import pandas as pd\nimport numpy as np\nimport pickle\nimport matplotlib.pyplot as plt\n\n'
    data_pickle = pickle.dumps(df)
    code_to_run = f"{import_line}\n{code}\n\ndf = pickle.loads({data_pickle!r})\nresult = {function_name}(df)\nprint(pickle.dumps(result))"

    return code_to_run

def extract_func_info(code, df):
    lines = code.split("\n")
    for line in lines:
        if is_assignment_statement(line):
            exec(line)

    pattern = re.compile(r'(\w+)\((.*?)\)')
    matches = pattern.findall(code)

    func_names = []
    func_args = []
    func_kwargs = []
    for match in matches:
        function_name = match[0]
        func_names.append(function_name)
        arguments = match[1]
        current_func_args = []
        current_func_kwargs = {}
        if len(arguments):
            arguments = arguments.split(',')
            arguments = [v.replace(' ', '') for v in arguments]
            for arg in arguments:
                # keyword arg
                if '=' in arg:
                    k, v = arg.split('=')
                    if is_subscript_and_index(v) or is_constant(v):
                        v = eval(v)
                    else:
                        v = v.replace("\'", '')
                    current_func_kwargs[k] = v
                else:
                    current_func_args.append(eval(arg))
        func_args.append(current_func_args)
        func_kwargs.append(current_func_kwargs)
    return func_names, func_args, func_kwargs


def decodestdoutput(output_from_std):
    output = ast.literal_eval(output_from_std)
    data = pickle.loads(output)
    return data


def is_assignment_statement(line):
    try:
        tree = ast.parse(line)
        for node in ast.walk(tree):
            if isinstance(node, ast.Assign):
                return True
    except SyntaxError:
        return False

    return False


def is_subscript_and_index(line):
    try:
        tree = ast.parse(line)
        subscript, index = False, False
        for node in ast.walk(tree):
            if isinstance(node, ast.Subscript):
                subscript = True
            elif isinstance(node, ast.Index):
                index = True
        return subscript and index
    except SyntaxError:
        return False


def is_constant(line):
    try:
        tree = ast.parse(line)
        for node in ast.walk(tree):
            if isinstance(node, ast.Constant):
                return True
    except SyntaxError:
        return False

    return False

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i + 2], 16) for i in (0, 2, 4))
