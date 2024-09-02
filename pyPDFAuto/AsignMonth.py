def AsignMonth (month):
    numMonth = 0
    if month == 'enero':
        numMonth = 24
    elif month == 'febrero':
        numMonth = 25
    elif month == 'marzo':
        numMonth = 26
    elif month == 'abril':
        numMonth = 27        
    elif month == 'mayo':
        numMonth = 28
    elif month == 'junio':
        numMonth = 29
    elif month == 'julio':
        numMonth = 30
    elif month == 'agosto':
        numMonth = 31
    elif month == 'septiembre':
        numMonth = 20
    elif month == 'octubre':
        numMonth = 21
    elif month == 'noviembre':
        numMonth = 22
    elif month == 'diciembre':
        numMonth = 23
    else:
        numMonth = 23
        
    return(int(numMonth))
    