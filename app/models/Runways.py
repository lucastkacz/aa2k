class CustomDict(dict):
    def __getattr__(self, attr):
        return self[attr]

    def __setattr__(self, attr, value):
        self[attr] = value


runway = CustomDict({"EZE_11_29": 3000, "EZE_17_35": 3300, "AEP_09_27": 2200})

# Access using dot notation
print(runway.EZE_11_29)  # 3000

# Access using key-value notation
print(runway["EZE_11_29"])  # 3000
