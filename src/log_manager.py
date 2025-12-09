import logging
import os
from logging.handlers import RotatingFileHandler


class LogManager:
    @staticmethod
    def get_logger(name="PaddleOCRVL", log_file="ocr_assistant.log"):
        """获取日志记录器

        Args:
            name: 日志记录器名称
            log_file: 日志文件名

        Returns:
            logging.Logger: 配置好的日志记录器
        """
        # 创建日志目录
        log_dir = "logs"
        os.makedirs(log_dir, exist_ok=True)
        log_path = os.path.join(log_dir, log_file)

        # 创建日志记录器
        logger = logging.getLogger(name)
        logger.setLevel(logging.DEBUG)

        # 避免重复添加处理器
        if not logger.handlers:
            # 创建控制台处理器
            console_handler = logging.StreamHandler()
            console_handler.setLevel(logging.DEBUG)

            # 创建文件处理器
            file_handler = RotatingFileHandler(
                log_path, maxBytes=10 * 1024 * 1024, backupCount=5, encoding='utf-8'
            )
            file_handler.setLevel(logging.INFO)

            # 创建日志格式
            formatter = logging.Formatter(
                '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
            )
            console_handler.setFormatter(formatter)
            file_handler.setFormatter(formatter)

            # 添加处理器
            logger.addHandler(console_handler)
            logger.addHandler(file_handler)

        return logger


# 创建全局日志记录器
logger = LogManager.get_logger()
