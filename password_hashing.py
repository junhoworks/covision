import bcrypt

# 암호 해싱
password = ""
hashed_password = bcrypt.hashpw(password.encode(), bcrypt.gensalt())

# 해싱된 암호 출력 (이 값을 데이터베이스에 저장)
print("해싱된 암호:", hashed_password)