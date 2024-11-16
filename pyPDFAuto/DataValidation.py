def PDF_DataValidationFull (Pdf_name, month, hours, courses, notes, star, end):
    if not isinstance(hours, str):
        return False, "Solo puede ingresar letras en el espacio de horas"
    elif not len(hours) == 1:
        return False, "Solo puede ingresar una letra en el espacio de horas, ingrese la letra que indique la columna de Excel donde se encuentren ubicadas las horas"
    elif not isinstance(Pdf_name, str):
        return False, "Solo puede ingresar letras en el espacio nombre"
    elif not len(Pdf_name) == 1:
        return False, "Solo puede ingresar una letra en el espacio de nombre, ingrese la letra que indique la columna de Excel donde se encuentren ubicadas los nombres"
    elif not isinstance(month, str):
        return False, "Solo puede ingresar letras en el espacio mes"
    elif not isinstance(courses, str):
        return False, "Solo puede ingresar letras en el espacio cursos"
    elif not len(courses) == 1:
        return False, "Solo puede ingresar una letra en el espacio de cursos, ingrese la letra que indique la columna de Excel donde se encuentren ubicadas los cursos"
    elif not isinstance(notes, str):
        return False, "Solo puede ingresar letras en el espacio notas"
    elif not len(notes) == 1:
        return False, "Solo puede ingresar una letra en el espacio de notas, ingrese la letra que indique la columna de Excel donde se encuentren ubicadas los notas"
    elif not isinstance(star, int):
        return False, "Solo puede ingresar numeros en el espacio inicio" 
    elif not isinstance(end, int):
        return False, "Solo puede ingresar numeros en el espacio fin"
    elif (Pdf_name == None or month == None or hours == None or courses == None or notes == None or star == None or end == None):
        return False, "Por favor rellene todos los espacios necesarios"
    else:
         return True


def PDF_DataValidationNormal (Pdf_name, month, PreachCheck, courses, notes, star, end):
    if not isinstance(Pdf_name, str):
        return False, "Solo puede ingresar letras en el espacio nombre"
    elif not len(Pdf_name) == 1:
        return False, "Solo puede ingresar una letra en el espacio de nombre, ingrese la letra que indique la columna de Excel donde se encuentren ubicadas los nombres"
    elif not isinstance(month, str):
        return False, "Solo puede ingresar letras en el espacio mes"
    elif not isinstance(courses, str):
        return False, "Solo puede ingresar letras en el espacio cursos"
    elif not len(courses) == 1:
        return False, "Solo puede ingresar una letra en el espacio de cursos, ingrese la letra que indique la columna de Excel donde se encuentren ubicadas los cursos"
    elif not isinstance(notes, str):
        return False, "Solo puede ingresar letras en el espacio notas"
    elif not len(notes) == 1:
        return False, "Solo puede ingresar una letra en el espacio de notas, ingrese la letra que indique la columna de Excel donde se encuentren ubicadas los notas"
    elif not isinstance(PreachCheck, str):
        return False, "Solo puede ingresar letras en el espacio predicacion"
    elif not len(PreachCheck) == 1:
        return False, "Solo puede ingresar una letra en el espacio de predicacion, ingrese la letra que indique la columna de Excel donde se encuentren ubicados los datos de predicacion"
    elif not isinstance(star, int):
        return False, "Solo puede ingresar numeros en el espacio inicio" 
    elif not isinstance(end, int):
        return False, "Solo puede ingresar numeros en el espacio fin" 
    elif (Pdf_name == None or month == None or PreachCheck == None or courses == None or notes == None or star == None or end == None):
        return False, "Por favor rellene todos los espacios necesarios"
    else:
         return True
    

    