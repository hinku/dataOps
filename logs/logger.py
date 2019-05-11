import logging
from logging.handlers import RotatingFileHandler
from logging import config
import yaml
import os

def setup_logging(default_path = 'logging.yaml',default_level = logging.INFO, env_key = 'LOG_CFG'):
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)),default_path)
    value = os.getenv(env_key,None)
    if value:
        path = value
    if os.path.exists(path):
        with open(path,"r") as f:
            cfg = yaml.load(f)
            config.dictConfig(cfg)
    else:
        logging.basicConfig(level = default_level)

    logger = logging.getLogger(__name__)

    return logger

def initLogger():
    logger = logging.getLogger(__name__)
    logger.setLevel(level = logging.INFO)
    #定义一个RotatingFileHandler，最多备份3个日志文件，每个日志文件最大1K
    rHandler = RotatingFileHandler("log.txt",maxBytes = 1*1024*1024,backupCount = 3)
    rHandler.setLevel(logging.INFO)
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    rHandler.setFormatter(formatter)
    
    console = logging.StreamHandler()
    console.setLevel(logging.INFO)
    console.setFormatter(formatter)
    
    logger.addHandler(rHandler)
    logger.addHandler(console)

    return logger



logger = setup_logging()

if __name__ == "__main__":
    setup_logging()
    logging.info('logger setup')