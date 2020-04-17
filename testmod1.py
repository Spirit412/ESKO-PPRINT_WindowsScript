import logging

def add(x, y):
    """"""
    logger = logging.getLogger('testmod1')
    logger.info("added %s and %s to get %s" % (x, y, x+y))
    return x+y
