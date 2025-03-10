# https://www.geeksforgeeks.org/python-oops-concepts/ 

"""
Types of Inheritance:

1. Single Inheritance: A child class inherits from a single parent class.
2. Multiple Inheritance: A child class inherits from more than one parent class.
3. Multilevel Inheritance: A child class inherits from a parent class, which in turn inherits from another class.
4. Hierarchical Inheritance: Multiple child classes inherit from a single parent class.
5. Hybrid Inheritance: A combination of two or more types of inheritance.
"""

# class Dog:
#     def __init__(self, name):
#         self.name = name

#     def display_name(self):
#         print(f"Dog's Name: {self.name}")

# class Labrador(Dog):  # Single Inheritance
#     def sound(self):
#         print("Labrador woofs")

# class GuideDog(Labrador):  # Multilevel Inheritance
#     def guide(self):
#         print(f"{self.name} guides the way!")

# class Friendly:
#     def greet(self):
#         print("Friendly!")

# class GoldenRetriever(Dog, Friendly):  # Multiple Inheritance
#     def sound(self):
#         print("Golden Retriever Barks")


# lab = Labrador("Buddy")
# lab.display_name()
# lab.sound()

# guide_dog = GuideDog("Max")
# guide_dog.display_name()
# guide_dog.guide()

# retriever = GoldenRetriever("Charlie")
# retriever.display_name()
# retriever.greet()
# retriever.sound()



### --------------------------------------------------------------------------------------------------


"""
Types of Encapsulation:

1. Public Members: Accessible from anywhere.
2. Protected Members: Accessible within the class and its subclasses.
3. Private Members: Accessible only within the class.
"""

# class Dog:
#     def __init__(self, name, breed, age):
#         self.name = name  # Public attribute
#         self._breed = breed  # Protected attribute '_'
#         self.__age = age  # Private attribute '__'

#     # Public method
#     def get_info(self):
#         return f"Name: {self.name}, Breed: {self._breed}, Age: {self.__age}"

#     # Getter and Setter for private attribute
#     def get_age(self):
#         return self.__age

#     def set_age(self, age):
#         if age > 0:
#             self.__age = age
#         else:
#             print("Invalid age!")


# dog = Dog("Buddy", "Labrador", 3)

# print(dog.name)  # Accessible
# print(dog._breed)  # Accessible but discouraged outside the class
# print(dog.get_age()) # private -> accessible through getter

# # Modifying private member using setter
# dog.set_age(5)
# print(dog.get_info())



### --------------------------------------------------------------------------------------------------


"""
Types of Abstraction:
(abstraction = hiding complex implementation details and providing a simplified interface)

1. Partial Abstraction: Abstract class contains both abstract and concrete methods.
2. Full Abstraction: Abstract class contains only abstract methods (like interfaces).
"""

from abc import ABC, abstractmethod

class Dog(ABC):  # Abstract Class
    def __init__(self, name):
        self.name = name

    @abstractmethod
    def sound(self):  # Abstract Method
        pass

    def display_name(self):  # Concrete Method
        print(f"Dog's Name: {self.name}")

class Labrador(Dog):  # Partial Abstraction
    def sound(self):
        print("Labrador Woof!")

class Beagle(Dog):  # Partial Abstraction
    def sound(self):
        print("Beagle Bark!")

# Example Usage
dogs = [Labrador("Buddy"), Beagle("Charlie")]
for dog in dogs:
    dog.display_name()  # Calls concrete method
    dog.sound()  # Calls implemented abstract method

tara = Dog()
print(tara)