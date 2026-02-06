import vnstock
print("Modules in vnstock:", dir(vnstock))

try:
  from vnstock import Vnstock
  print("Found Vnstock class")
except:
  print("No Vnstock class")
