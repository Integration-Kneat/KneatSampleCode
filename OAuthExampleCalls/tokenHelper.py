
filename = "bearerToken.cfg"


def getToken():

    with open(filename) as f:
        bearerToken = f.read()
        f.close()

    return  bearerToken


def saveToken(bearerToken):

    file = open(filename, 'w')
    file.write(bearerToken)
    file.close()