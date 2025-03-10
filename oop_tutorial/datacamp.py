# https://www.datacamp.com/tutorial/python-oop-tutorial

class Dog:
    
    def __init__(self, name, age):
        self.name = name
        self.age = age
        
    def bark(self):
        print("waf waf!")
        
    def doginfo(self):
        print(f"Deze hond noemt {self.name} en is {self.age} jaar.")
        
    def birthday(self):
        self.age += 1
        print(f"It's {self.name}'s birthday!")
        
    def setBuddy(self, buddy):
        self.buddy = buddy
        buddy.buddy = self
        print(f"{self.name} and {buddy.name} are buddies!")

    
cuzco = Dog("Cuzco", 4)
tara = Dog("Tara", 13)
spot = Dog("Spot", 8)

cuzco.doginfo()

tara.age = 15
tara.doginfo()

spot.doginfo()
spot.birthday()
spot.doginfo()

spot.setBuddy(tara)
tara.buddy.doginfo()
spot.buddy.doginfo()
