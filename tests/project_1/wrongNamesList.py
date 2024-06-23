
def wrongWorkbookNames():
    return ["wrongNamesList", "wrongNamesList2", "wrongNamesList3"]

def wrongSheetNames():

    return ["wrongNamesList ", 
            " wrongNamesList2", 
            "", 
            "'wrongNamesList3'",
            "[wrongNamesList4]",
            "\{wrongNamesList5\}"]

VALID_JSON_DATA =  """
            {
                "sheets":[
                    {
                        "name":"Sheet1",
                        "cell-contents":{
                            "A1":"'123",
                            "B1":"5.3",
                            "C1":"=A1*B1"
                        }
                    },
                    {
                        "name":"Sheet2",
                        "cell-contents":{}
                    }
                ]
            }
        """