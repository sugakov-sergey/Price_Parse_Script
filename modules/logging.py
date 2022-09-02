def logging(func):
    """ Decorator for logging script functions """
    def log_wrapp(*args, **kwargs):
        import logging
        logging.basicConfig(
            level=logging.DEBUG,
            filename="logs/parse_log.log",
            format="%(asctime)s - %(module)s - %(levelname)s - %(funcName)s: %(lineno)d - %(message)s",
            datefmt='%H:%M:%S',
            filemode='w', )
        msg = f'{func.__name__} started\n {args}\n {kwargs}'
        logging.info(msg)
        func(*args, **kwargs)
        msg = f'{func.__name__} finished'
        logging.info(msg)
    return log_wrapp
