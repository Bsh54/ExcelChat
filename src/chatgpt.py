import openai


class ChatBot(object):
    def __init__(self):
        self.api_key = None
        self.api_base = "https://shads229-personnal-ai.hf.space/v1"
        self.model = 'deepseek-chat'
        openai.api_base = self.api_base
        openai.api_key = None
        self.history = []

    def set_api_key(self, api_key):
        self.api_key = api_key
        openai.api_key = api_key

    def clear_history(self):
        self.history = []

    def get_response(self, content, system_prompt=None):
        # Mettre à jour ou ajouter le prompt système (contexte du DataFrame)
        if system_prompt:
            # Si le premier message est un système, on le met à jour, sinon on l'insère
            if self.history and self.history[0]['role'] == 'system':
                self.history[0]['content'] = system_prompt
            else:
                self.history.insert(0, {'role': 'system', 'content': system_prompt})

        self.history.append({'role': 'user', 'content': content})

        if self.api_key is not None:
            openai.api_base = self.api_base
            openai.api_key = self.api_key
            response = openai.ChatCompletion.create(
                model=self.model,
                messages=self.history,
                temperature=0.3)

            message = response.choices[0]['message']['content']
            token_count = response['usage']['total_tokens']

            self.history.append({'role': 'assistant', 'content': message})

            return message, int(token_count)
        else:
            return '', 0
