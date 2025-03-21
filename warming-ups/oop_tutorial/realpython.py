# https://realpython.com/python3-object-oriented-programming/ 

class Dog:
    # class attributes: (same value for all class instances)
    species = "Canis familiaris"
    
    def __init__(self, name, age):
        # instance attributes:
        self.name = name
        self.age = age
        
    # def description(self):
    #     return f"{self.name} is {self.age} years old"
    def __str__(self):
        return f"{self.name} is {self.age} years old"


    def speak(self, sound):
        return f"{self.name} barks: {sound}"
    
class JackRussellTerrier(Dog):
    # # overwriting parent .speak() method
    # def speak(self, sound="Arf"):
    #     return f"{self.name} says {sound}"
    
    # not losing any changes in the parent .speak() method
    def speak(self, sound="Arf"):
        return super().speak(sound) # accessing parent class through super()

class Dachshund(Dog):
    pass

class Bulldog(Dog):
    pass
        
        
miles = JackRussellTerrier("Miles", 4)
buddy = Dachshund("Buddy", 9)
jack = Bulldog("Jack", 3)
jim = Bulldog("Jim", 5)

# print(miles)
# print(type(miles))
# print(isinstance(miles, Bulldog))

print(miles.speak())
print(miles.speak("Grrr"))
print(jim.speak("Woof"))




### --------------------------------------------------------------------------------------------------


class Parent:
    hair_color = "brown"
    speaks = ["English"]

class Child(Parent):
    def __init__(self):
        super().__init__() # inherit attributes from parent
        self.speaks.append("German")
        
        
        
"""
Note: In the above examples, the class hierarchy is very straightforward. The JackRussellTerrier class has a single parent class, Dog. In real-world examples, the class hierarchy can get quite complicated.

The super() function does much more than just search the parent class for a method or an attribute. It traverses the entire class hierarchy for a matching method or attribute. If you aren’t careful, super() can have surprising results.
"""