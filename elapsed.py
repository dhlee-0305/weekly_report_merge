import time

from logger import *

log = getLogger('elapsed')

def elapsed(original_func):
    def _elapsed(*args, **kwargs):
        startTime  = time.time()
        result = original_func(*args, **kwargs)
        endTime = time.time()
        log.debug(original_func.__name__+"() 수행시간: %f 초" % (endTime - startTime))
        return result
    return _elapsed
