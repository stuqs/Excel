import time


#--------------------------------------------------------------------------------------------------------------------

def timer(label='', trace=True):
    class Timer:
        def __init__(self, func):
            self.func = func
            self.alltime = 0
        def __call__(self, *args, **kargs):
            start = time.clock()
            result = self.func(*args, **kargs)
            elapsed = time.clock() - start
            self.alltime += elapsed
            if trace:
                format = '%s %s: %.5f, %.5f'
                values = (label, self.func.__name__, elapsed, self.alltime)
                print (elapsed)
                # print(format % values)
                return result
    return Timer

#--------------------------------------------------------------------------------------------------------------------

class singleton:
    def __init__(self, aClass):
        self.aClass = aClass
        self.instance = None
    def __call__(self, *args, **kargs):
        if self.instance is None:
            self.instance = self.aClass(*args, **kargs)
        return self.instance

# def singleton(aClass):
#     instance = None
#     def onCall(*args, **kargs):
#         nonlocal instance
#         if instance == None:
#             instance = aClass(*args, **kargs)
#         return instance
#     return onCall

#--------------------------------------------------------------------------------------------------------------------
# Decorators Private and Public

traceMe = False
def trace(*args):
    if traceMe:
        print('[' + ' '.join(map(str, args)) + ']')
def accessControl(failIf):
    def onDecorator(aClass):
        class onInstance:
            def __init__(self, *args, **kargs):
                self.__wrapped = aClass(*args, **kargs)
            def __getattr__(self, attr):
                trace('get:', attr)
                if failIf(attr):
                    raise TypeError('private attribute fetch: ' + attr)
                else:
                    return getattr(self.__wrapped, attr)
            def __setattr__(self, attr, value):
                trace('set:', attr, value)
                if attr == '_onInstance__wrapped' :
                    self.__dict__[attr] = value
                elif failIf(attr):
                    raise TypeError('private attribute change: ' + attr)
                else:
                    setattr(self.__wrapped, attr, value)
        return onInstance
    return onDecorator
def Private(*attributes):
    return accessControl(failIf=(lambda attr: attr in attributes))
def Public(*attributes):
    return accessControl(failIf=(lambda attr: attr not in attributes))

#--------------------------------------------------------------------------------------------------------------------


