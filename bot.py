"""
Please make sure you install the bot dependencies with `pip install --upgrade -r requirements.txt`
in order to get all the dependencies on your Python environment.
"""
from botcity.maestro import *
from functions import *

BotMaestroSDK.RAISE_NOT_CONNECTED = False

def main():
    maestro = BotMaestroSDK.from_sys_args()
    execution = maestro.get_execution()

    setup_logging()
    logging.info("Inicio - Atividade 3 - Python & Automation Anywhere")

    if execution.task_id == 0:
        logging.info("Maestro desativado -> Executando localmente")
        maestro = None

    if maestro:
        print(f"Task ID is: {execution.task_id}")
        print(f"Task Parameters are: {execution.parameters}")

    url = "https://rpachallenge.com/"

    web_bot = bot_driver_setup()

    download_phase = download_excel(web_bot, url)
    if not download_phase:
        logging.info("Ocorreu um erro na fase de download.")
    else:
        logging.info("Download concluido com sucesso.")

    data = read_excel()

    start_phase = start_challenge(web_bot)
    if not start_phase:
        logging.info("Ocorreu um erro na fase de clicar no botao de start.")
    else:
        logging.info("Foi iniciado o processo.")

    fill_phase = fill_fields(web_bot, data)
    if not fill_phase:
        logging.info("Ocorreu um erro na fase de preencher os campos.")
    else:
        logging.info("Elementos preenchidos corretamente.")
        success_message(web_bot)

    browse_close(web_bot)

    logging.info("Fim - Atividade 3 - Python & Automation Anywhere")

    if maestro:
        maestro.finish_task(
            task_id=execution.task_id,
            status=AutomationTaskFinishStatus.SUCCESS,
            message="Task Finished OK.",
            total_items=0,
            processed_items=0,
            failed_items=0
        )

def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()
