import logging

def get_logger():
    logger = logging.getLogger("api_logger")
    logger.setLevel(logging.DEBUG)

    # Console Handler (DEBUG+)
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.DEBUG)

    # File Handler (INFO+)
    file_handler = logging.FileHandler("api.log")
    file_handler.setLevel(logging.INFO)

    # Formato
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(lineno)d - %(levelname)s - %(message)s')
    console_handler.setFormatter(formatter)
    file_handler.setFormatter(formatter)

    # Adiciona os handlers
    logger.addHandler(console_handler)
    logger.addHandler(file_handler)

    return logger
