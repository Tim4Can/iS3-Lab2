from config import tasks,projects,datatypes
import os
import sys
sys.path.append("./library")
import FileProcessBasic



inputPath = 'D:/Death in TJU/Junior_2nd/iS3 Lab2/Tasks/task2/GPR'
outputPath= ''



# extract key word
def extract(filename):
	project=""
	datatype=""

	for prj in projects:
		if(filename.find(prj)>=0):
			project=prj;
			break;

	for dt in datatypes:
		if(filename.find(dt)>=0):
			datatype=dt;
			break;

	if(project=="" or datatype==""):
		return "null"
	else:
		return project+datatype



if  __name__=="__main__":

	# input 
	files=os.listdir(inputPath)

	# traverse files
	for file in files:
		# get dict key
		name=extract(file)
		#print(file+"\t"+name)

		if name in tasks:

			# import module
			module = __import__("library."+tasks[name][0],fromlist=True) 
			if hasattr(module, tasks[name][1]):
				# import certain class
				cn= getattr(module, tasks[name][1])
				# invoke function 'run'
				func=getattr(cn(),'run') 
				func(inputPath,outputPath)
		else:
			print("404")

