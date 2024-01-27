# coding: utf-8
from __future__ import unicode_literals

import os
import uuid

import requests
import json
import uno
import unohelper

import pyperclip

from com.sun.star.awt.MessageBoxType import INFOBOX, QUERYBOX
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK, BUTTONS_OK_CANCEL
from com.sun.star.awt import XActionListener

component_ctx = uno.getComponentContext()
service_manager = component_ctx.getServiceManager()

SOME_ERROR = """Произошла ошибка во время обращения к api GigaChat.\n Error:\n {response}"""
AUTH_ERROR = """Произошла ошибка во время авторизации. Пожалуйста, проверьте ваш токен.\n Error:\n {response}"""

TEXT_INFO = """
Это плагин, который расширяет функциональность LibreOffice. Он имеет большой функционал работы с текстом. Для использования достаточно выделить текст и выбрать нужную команду.
Плагин использует нейронную сеть GigaChat от Сбербанка.

Приятного пользования.
"""

PROMPT_FINISH = "Ты система продолжения текста. Система продолжения текста работает следующим образом: когда ей " \
                "предоставляется начальный текст, она анализирует его стилистику и контекст, а затем генерирует " \
                "продолжение текста в том же стиле и с учетом данного контекста. Ответ системы всегда содержит только " \
                "продолжение текста, без дополнительных комментариев или информации. \nПродолжи текст: {text}"

PROMPT_FIX = """Ты система исправления ошибок в тексте.
Ты умеешь только исправить текст, ничего больше. 
Если не можешь найти ошибки, ты должна вернуть тот же текст.
Не добавляй знак " в начале и в конце текста, если их не было.
Исправь грамматические и пунктуационные ошибки, в тексте:
{text}
"""

PROMPT_MAIN_THEMES = """Ты - система выявление основных идей текста.
Когда тебе приходит текст, ты находишь главные идеи текста.
В своем ответе ты пишешь только идеи, которые нашла. Больше ничего.
Какие основные идеи в тексте:  
{text}"""

PROMPT_TO_CONVERSATIONAL = """Ты - система, которая пишет текст только в разговорном стиле, ничего больше.
Когда тебе приходит текст, ты пытаешься переписать его в разговорном стиле.
Если не можешь поменять стиль текста, ты должна вернуть тот же текст.
Перепиши текст в разговорном стиль:
{text}"""

PROMPT_TO_OFFICIAL = """Ты - система, которая пишет текст только в деловом стиле, ничего больше.
Когда тебе приходит текст, ты пытаешься переписать его в деловом стиле.
Если не можешь поменять стиль текста, ты должна вернуть тот же текст.
Перепиши текст в деловом стиль:
{text}"""

PROMPT_EXPLANATIONS = """Ты - система, которая объясняет текст и термины в нем.
Пиши только объяснения текста. Ничего больше.
Вот текст: 
{text}"""

PROMPT_TO_SIMPLE = """Ты - система, которая упрощает текст.
Ты должна улучшить читабельность и понятность текста.
Если не можешь изменить текст, ты должна вернуть тот же текст.
Вот текст: 
{text}"""

PROMPT_TO_OPTIONS_OF_CONTENT = """Ты - система подбора оглавления.
Тебе приходит текст и ты подбираешь 3 наиболее подходящих вариантов оглавления к тексту
Вот текст: 
{text}"""

PROMPT_MAIN_THEMES_KEY = "PROMPT_MAIN_THEMES"
PROMPT_PROMPT_FIX_KEY = "PROMPT_FIX"
PROMPT_PROMPT_FINISH_KEY = "PROMPT_FINISH"
PROMPT_TO_CONVERSATIONAL_KEY = "PPROMPT_TO_CONVERSATIONAL"
PROMPT_TO_SIMPLE_KEY = "PROMPT_TO_SIMPLE"
PROMPT_EXPLANATIONS_KEY = "PROMPT_EXPLANATIONS"
PROMPT_TO_OFFICIAL_KEY = "PROMPT_TO_OFFICIAL"
PROMPT_TO_OPTIONS_OF_CONTENT_KEY = "PROMPT_TO_OPTIONS_OF_CONTENT"

CONFIG_FILE = "config_enterprise.json"

BASE_CONFIG = {
    "token": "YOUR_TOKEN",
    "version": "0.0.1",
    "update": True,
    PROMPT_MAIN_THEMES_KEY: PROMPT_MAIN_THEMES,
    PROMPT_PROMPT_FIX_KEY: PROMPT_FIX,
    PROMPT_PROMPT_FINISH_KEY: PROMPT_FINISH,
    PROMPT_TO_CONVERSATIONAL_KEY: PROMPT_TO_CONVERSATIONAL,
    PROMPT_TO_OFFICIAL_KEY: PROMPT_TO_OFFICIAL,
    PROMPT_EXPLANATIONS_KEY: PROMPT_EXPLANATIONS,
    PROMPT_TO_SIMPLE_KEY: PROMPT_TO_SIMPLE,
    PROMPT_TO_OPTIONS_OF_CONTENT_KEY: PROMPT_TO_OPTIONS_OF_CONTENT,
}

PROMPT_MAIN_THEMES = lambda: get_config(PROMPT_MAIN_THEMES_KEY)
PROMPT_FIX = lambda: get_config(PROMPT_PROMPT_FIX_KEY)
PROMPT_FINISH = lambda: get_config(PROMPT_PROMPT_FINISH_KEY)
PROMPT_TO_CONVERSATIONAL = lambda: get_config(PROMPT_TO_CONVERSATIONAL_KEY)
PROMPT_TO_OFFICIAL = lambda: get_config(PROMPT_TO_OFFICIAL_KEY)
PROMPT_EXPLANATIONS = lambda: get_config(PROMPT_EXPLANATIONS_KEY)
PROMPT_TO_SIMPLE = lambda: get_config(PROMPT_TO_SIMPLE_KEY)
PROMPT_TO_OPTIONS_OF_CONTENT = lambda: get_config(PROMPT_TO_OPTIONS_OF_CONTENT_KEY)

response_function = lambda text_range, text: None
response_text = ""

folder_path = os.path.join(os.path.expanduser("~"), ".gigachat")
if not os.path.exists(folder_path):
    # Создаем новую папку
    os.makedirs(folder_path)

CONFIG_PATH = os.path.join(folder_path, CONFIG_FILE)


def get_text_range():
    return get_selection().getByIndex(0)


def get_selection():
    return get_model().getCurrentController().getSelection()


def get_model():
    document = service_manager.createInstanceWithContext("com.sun.star.frame.Desktop", component_ctx)
    return document.getCurrentComponent()


def replace_selection_text(text_range, text: str):
    text_range.setString(text)


def insert_before_selection_text(text_range, text: str):
    text_range.setString(text + text_range.getString())


def insert_after_selection_text(text_range, text: str):
    text_range.setString(text_range.getString() + text)


def get_config(key=None):
    try:
        with open(CONFIG_PATH, 'r', encoding="utf-8") as json_file:
            data = json.load(json_file)
        if not key:
            return data
        prompt = data.get(key)
        if prompt is None:
            prompt = BASE_CONFIG.get(key)
        return prompt
    except FileNotFoundError:
        create_config()
        if not key:
            return BASE_CONFIG
        return BASE_CONFIG.get(key)


def create_config(save_config: dict = None):
    if save_config is None:
        save_config = BASE_CONFIG
    with open(CONFIG_PATH, 'w', encoding="utf-8") as json_file:
        json.dump(save_config, json_file, indent=4, ensure_ascii=False)


def check_config():
    config = get_config()
    rewrite = False
    if config.get("version", "") != BASE_CONFIG.get("version") and config.get("update", True):
        rewrite = True
        for prompt_name, prompt in BASE_CONFIG.items():
            if prompt_name == "token":
                continue
            config[prompt_name] = prompt

    for prompt_name, prompt in BASE_CONFIG.items():
        if config.get(prompt_name) is None:
            config[prompt_name] = prompt
            rewrite = True

    if rewrite:
        create_config(config)


def update_config(key, val):
    with open(CONFIG_PATH, 'r', encoding="utf-8") as json_file:
        data = json.load(json_file)

    data[key] = val
    with open(CONFIG_PATH, 'w', encoding="utf-8") as json_file:
        json.dump(data, json_file, indent=4, ensure_ascii=False)


def get_select_text_with_range():
    text_range = get_text_range()
    selected_text = text_range.getString()
    return selected_text, text_range


def get_token():
    config = get_config()
    if config.get('token') == BASE_CONFIG.get("token"):
        get_msg_box("Пожалуйста, добавьте ваш токен от API GigaChat")
        return add_token_dialog()
    return config.get("token")


class TokenListener(unohelper.Base, XActionListener):
    def __init__(self, token, dialog):
        self.dialog = dialog
        self.token = token

    def actionPerformed(self, action_event):
        source = action_event.Source
        if source.Model.Name == "Ok":
            token = self.token.Text
            update_config("token", token)
            self.dialog.endExecute()
        elif source.Model.Name == "Cancel":
            self.dialog.endExecute()


def add_token_dialog():
    dp = service_manager.createInstanceWithContext("com.sun.star.awt.DialogProvider2", component_ctx)
    dialog = dp.createDialog("vnd.sun.star.extension://gigachatenterprise/dialogs/AddTokenDialog.xdl")
    dialog.setTitle("Добавление токена от API GigaChat")
    ok = dialog.getControl("Ok")
    cancel = dialog.getControl("Cancel")
    token = dialog.getControl("Token")
    token.setText(get_config("token"))
    ok.addActionListener(TokenListener(token, dialog))
    cancel.addActionListener(TokenListener(token, dialog))

    dialog.execute()


class ResponseListener(unohelper.Base, XActionListener):
    BUTTON_TO_FUNCTION = {
        "before": insert_before_selection_text,
        "after": insert_after_selection_text,
        "replace": replace_selection_text,
    }

    def __init__(self, text, text_range, dialog):
        self.dialog = dialog
        self.text = text
        self.text_range = text_range

    def actionPerformed(self, action_event):
        global response_function, response_text
        btn_name = action_event.Source.Model.Name
        if btn_name == "copy":
            pyperclip.copy(self.text.getText())
            get_msg_box("скопировано")
        response_function = ResponseListener.BUTTON_TO_FUNCTION[btn_name]
        response_text = self.text.getText()
        self.dialog.endExecute()


def get_response_dialog(text_to_insert, text_range, title: str = "Ответ от GigaChat"):
    if not text_to_insert:
        return
    dp = service_manager.createInstanceWithContext("com.sun.star.awt.DialogProvider2", component_ctx)
    dialog = dp.createDialog("vnd.sun.star.extension://gigachatenterprise/dialogs/ResponseDialog.xdl")
    dialog.setTitle(title)
    text = dialog.getControl("text")
    before_button = dialog.getControl("before")
    after_button = dialog.getControl("after")
    replace_button = dialog.getControl("replace")
    copy_button = dialog.getControl("copy")
    text.setText(text_to_insert)
    replace_button.addActionListener(ResponseListener(text, text_range, dialog))
    copy_button.addActionListener(ResponseListener(text, text_range, dialog))
    after_button.addActionListener(ResponseListener(text, text_range, dialog))
    before_button.addActionListener(ResponseListener(text, text_range, dialog))

    dialog.execute()


class CustomPromptListener(unohelper.Base, XActionListener):
    def __init__(self, text, text_range, dialog):
        self.dialog = dialog
        self.text = text
        self.text_range = text_range

    def actionPerformed(self, action_event):
        global response_function, response_text
        btn_name = action_event.Source.Model.Name
        if btn_name == "paste":
            self.text.setText(self.text.getText() + self.text_range.getString())
        elif btn_name == "send":
            self.dialog.endExecute()
            text = get_info(self.text.getText())
            get_response_dialog(text, self.text_range)


def get_custom_prompt_dialog(text_range, title: str = "Ответ от GigaChat"):
    dp = service_manager.createInstanceWithContext("com.sun.star.awt.DialogProvider2", component_ctx)
    dialog = dp.createDialog("vnd.sun.star.extension://gigachatenterprise/dialogs/CustomPromptDialog.xdl")
    dialog.setTitle(title)
    text = dialog.getControl("text")
    send = dialog.getControl("send")
    paste = dialog.getControl("paste")
    send.addActionListener(CustomPromptListener(text, text_range, dialog))
    paste.addActionListener(CustomPromptListener(text, text_range, dialog))
    dialog.execute()


def get_msg_box(message):
    title = "Giga Chat Плагин"
    message_box_type = INFOBOX
    toolkit = service_manager.createInstanceWithContext("com.sun.star.awt.Toolkit", component_ctx)
    parent_win = toolkit.getDesktopWindow()
    msg_box = toolkit.createMessageBox(
        parent_win,
        message_box_type,
        BUTTONS_OK,
        title,
        message
    )
    msg_box.execute()


def show_error_message(message):
    if not message:
        return None
    title = "Ошибка с Giga Chat Плагин"
    toolkit = service_manager.createInstanceWithContext("com.sun.star.awt.Toolkit", component_ctx)
    parent_win = toolkit.getDesktopWindow()
    # Create a message box
    msgbox = toolkit.createMessageBox(
        parent_win,
        1,
        BUTTONS_OK,
        title,
        message
    )

    # Display the message box
    msgbox.execute()


def get_info(content: str = None):
    """Method sends POST request to API and returns result of working model.
    If content is None or empty it will show info about macros.

    Args:
        content: str - text data from office file - content for request.

    Returns:
        str - text from response to API.
    """
    if content is None or content.strip() == "":
        get_msg_box(TEXT_INFO)
        get_token()
        return ""
    token = get_token()
    if token == BASE_CONFIG.get(token):
        token = get_token()
    url_get_token = 'https://ngw.devices.sberbank.ru:9443/api/v2/oauth'
    data = {'scope': 'GIGACHAT_API_PERS'}
    headers_token = {
            'Authorization': f'Base {token}',
            'RqUID': str(uuid.uuid4())}
    token_response = requests.post(url=url_get_token, data=data, headers=headers_token, verify=False)
    if token_response.status_code != 200:
        show_error_message(AUTH_ERROR.format(response=token_response.text))
        add_token_dialog()
        return None
    access_token = token_response.json()["access_token"]
    url = "https://gigachat.devices.sberbank.ru/api/v1/chat/completions"
    payload = {
        "messages": [
            {
                "role": "system",
                "content": content,
            }
        ],
        "model": "GigaChat:latest",
        "temperature": 0.2
    }
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    response = requests.post(url, headers=headers, json=payload, verify=False)
    if response.status_code != 200:
        show_error_message(SOME_ERROR.format(response=response.text))
        return None
    return json.loads(response.text)["choices"][0]["message"]["content"]


def generate_command(prompt, title):
    def command():
        try:
            global response_function, response_text
            selected_text, text_range = get_select_text_with_range()
            if not selected_text:
                get_msg_box("Пожалуйста, выделите текст, который нужно обработать")
                return
            text = get_info(prompt().format(text=selected_text))
            response_function = lambda text_range, text: None
            response_text = ""
            get_response_dialog(text, text_range, title)
            response_function(text_range, response_text)

        except Exception as error:
            get_msg_box(f"Произошла ошибка: {error}")

    return command


def get_by_custom_prompt():
    global response_function, response_text
    selected_text, text_range = get_select_text_with_range()
    get_custom_prompt_dialog(text_range, "Укажите ваш запрос")
    response_function(text_range, response_text)


fix_text = generate_command(PROMPT_FIX, "Исправление ошибок")
text_to_conversational = generate_command(PROMPT_TO_CONVERSATIONAL, "В разговорном стиле")
text_to_official = generate_command(PROMPT_TO_OFFICIAL, "В деловом стиле")
text_to_simple = generate_command(PROMPT_TO_SIMPLE, "Упрощенный текст")
get_explanations = generate_command(PROMPT_EXPLANATIONS, "Объясненный текст")
get_options_of_content = generate_command(PROMPT_TO_OPTIONS_OF_CONTENT, "Оглавления")
continue_text = generate_command(PROMPT_FINISH, "Продолжение текста")
get_main_themes = generate_command(PROMPT_MAIN_THEMES, "Основные идеи")

check_config()

g_exportedScripts = (
    get_info,
    fix_text,
    get_main_themes,
    continue_text,
    add_token_dialog,
    text_to_official,
    text_to_conversational,
    text_to_simple,
    get_explanations,
    get_options_of_content,
    get_by_custom_prompt
)
