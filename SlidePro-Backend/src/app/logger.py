import logging
import datetime


def setup_logging():
    logging.basicConfig(filename=f'logs/{datetime.datetime.now().strftime("%Y-%m-%d")}.log',
                        filemode='a',
                        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                        level=logging.WARNING)
