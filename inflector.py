from pytrovich.detector import PetrovichGenderDetector
from pytrovich.enums import NamePart, Gender, Case
from pytrovich.maker import PetrovichDeclinationMaker

def inflect(dictionary):
    detector = PetrovichGenderDetector()
    try:
        gender = detector.detect(firstname=dictionary["BBBB"], middlename=dictionary["CCCC"])
    except:
        gender = Gender.MALE # если не удалось определить пол, укажем его как мужской

    maker = PetrovichDeclinationMaker()
    dictionary["AA"] = maker.make(NamePart.LASTNAME, gender, Case.GENITIVE, dictionary["AAAA"]) # склонение фамилии
    dictionary["BB"] = maker.make(NamePart.FIRSTNAME, gender, Case.GENITIVE, dictionary["BBBB"]) # склонение имени
    dictionary["CC"] = maker.make(NamePart.MIDDLENAME, gender, Case.GENITIVE, dictionary["CCCC"]) # склонение отчества
    # поместим ФИО в именительном и родительном падежах отдельно друг от друга
    # это пригодится, если в шаблоне ФИО встречается и в одной форме, и в другой

    return dictionary