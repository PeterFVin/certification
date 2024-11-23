from num_to_rus import Converter
import pymorphy3
from pytrovich.enums import NamePart, Gender, Case
from pytrovich.maker import PetrovichDeclinationMaker

conv = Converter()


def genitive(word):
    morph = pymorphy3.MorphAnalyzer()
    parsing_word = morph.parse(word)[0]
    gent = parsing_word.inflect({'gent'})
    return gent.word

choose_parameters = {
    'М': 'MALE',
    'Ж': 'FEMALE',
}


def genitive_name(choose_name, choose_gender, name):
    """
    Функция склоняет фамилию, имя или отчество в родительном падеже.
    
    Аргумент choose_name - FIRSTNAME, MIDDLENAME или LASTNAME (фамилия).
    Аргумент choose_sex - MALE или FEMALE, choose_gender - М или Ж,
    в словаре choose_parameters переводится в MALE или FEMALE.
    Пример <<чистого>> метода:
    maker.make(NamePart.FIRSTNAME, Gender.FEMALE, Case.GENITIVE, "Арина")
    """
    maker = PetrovichDeclinationMaker()
    name_part = getattr(NamePart, choose_name)
    gender = getattr(Gender, choose_parameters[choose_gender])
    return maker.make(name_part, gender, Case.GENITIVE, name)


def number_to_words(input_number):
    output_number = input_number.replace(' ', '')
    return conv.convert(int(output_number)).capitalize()


if __name__ == "__main__":
    # print(genitive('апрель'))
    print(genitive_name('LASTNAME', 'М', 'Павлов'))
    # print(number_to_words('2 000'))
