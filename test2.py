i = 1
lines = []
while i <= 5:
    exec("line{} = {}".format(i, i))
    i += 1
print(line4)
