import random

def generate_random_string(length):
    characters = '0123456789ABCDEFGHJKMNPQRSTVWXYZabcdefghjkmnpqrstvwxyz'
    random_string = ''.join(random.choices(characters, k=length))
    return random_string

def calculate_string_sum(string):
    sum = 0
    for char in string:
        if char.isdigit():
            sum += int(char)
    return sum
