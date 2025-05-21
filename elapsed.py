import time

from logger import *

log = getLogger('elapsed')

def elapsed(original_func):
    """
    함수 수행시간을 측정하는 데코레이터
    
    :param original_func: 원본 함수
    :return: _elapsed: 수행시간을 측정하는 함수
    """
    def _elapsed(*args, **kwargs):
        startTime  = time.time()
        result = original_func(*args, **kwargs)
        endTime = time.time()
        log.debug(original_func.__name__+"() 수행시간: %f 초" % (endTime - startTime))
        return result
    return _elapsed
