import random

def number_guessing_game():
    print("숫자 맞추기 게임에 오신 것을 환영합니다!")
    print("1부터 100 사이의 숫자를 맞춰보세요.")
    
    # 랜덤 숫자 생성
    secret_number = random.randint(1, 100)
    attempts = 0
    
    while True:
        try:
            guess = int(input("\n숫자를 입력하세요 (1-100): "))
            attempts += 1
            
            if guess < 1 or guess > 100:
                print("1에서 100 사이의 숫자를 입력해주세요!")
                continue
            
            if guess == secret_number:
                print(f"축하합니다! 정답입니다! 🎉")
                print(f"시도 횟수: {attempts}번")
                break
            elif guess < secret_number:
                print("더 큰 숫자입니다! ⬆️")
            else:
                print("더 작은 숫자입니다! ⬇️")
                
        except ValueError:
            print("올바른 숫자를 입력해주세요!")
    
    # 게임 재시작 여부 확인
    play_again = input("\n게임을 다시 하시겠습니까? (y/n): ").lower()
    if play_again == 'y':
        number_guessing_game()
    else:
        print("게임을 종료합니다. 감사합니다!")

if __name__ == "__main__":
    number_guessing_game()
