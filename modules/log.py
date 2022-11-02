import logging

logging.basicConfig(
    level=logging.DEBUG,
    filename="logs/parse_log.logs",
    format="%(asctime)s - line %(lineno)d ... %(funcName)s ... %(message)s",
    datefmt='%H:%M:%S',
    filemode='w', )

