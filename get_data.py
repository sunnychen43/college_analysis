import db
import sys

# url, file_name, school, year, type
if __name__ == '__main__':
	if(len(sys.argv) != 6):
		print('Invalid arguments')
		quit()
	db.create_db(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4], sys.argv[5])
