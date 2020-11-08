import logging
import os

logger = logging.getLogger()
logger.setLevel("INFO")
BASIC_FORMAT = '%(asctime)s - %(filename)s[line:%(lineno)d] - %(levelname)s: %(message)s'
DATE_FORMAT = '%Y-%m-%d %H:%M:%S'
formatter = logging.Formatter(BASIC_FORMAT, DATE_FORMAT)

chlr = logging.StreamHandler()
chlr.setFormatter(formatter)
# chlr.setLevel("INFO")


fhlr = logging.FileHandler('log.log', encoding='utf-8')

fhlr.setFormatter(formatter)

logger.addHandler(chlr)
logger.addHandler(fhlr)
logger.info(" ===== Begin log Info ===== ")
