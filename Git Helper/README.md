# 🖥 Git Helper

<img src="https://github.com/glowthem/ToyProject/blob/master/Git%20%EC%86%8C%EC%8A%A4%20%EC%97%85%EB%A1%9C%EB%93%9C%20%EC%9E%90%EB%8F%99%ED%99%94/imgSrc/Main.png" width="90%" height="90%" alt="메인화면">
<br/>

## ✔️ 프로젝트 진행기간
> 3학년 겨울 (2020년 1월~2월, 1개월)

## ✔️ 설계 방식
> 1. Visual Basic 활용, 지정한 경로 내의 모든 VBA 프로젝트에서 모듈, 클래스, 폼, 시트를 분리
> 2. 기업 서버 내의 GitLab repository 생성<br/>
> 3. git clone 수행
> 4. clone하여 생성된 로컬 저장소에 분리한 파일들을 옮긴 후 git add, commit, push까지 수행<br/>
> 5. 각 단계의 실행 결과를 실시간으로 엑셀 파일 각 탭에 출력

## ✔️ 기술 스택
> Git<br/>
> VBA(Visual Basic Application)<br/>
> MS Excel<br/>
> GitLab REST API

<br/>

## ✔️ 진행 방식
<br/>

> 1. 메인 화면의 소스 분할을 클릭하면 폴더를 선택할 수 있다. 지정한 경로 내의 모든 엑셀파일에서 모듈, 클래스, 폼, 시트를 분리한다.<br/>
> 각 파일경로, 파일 이름, 파일 번호, 프로그램 이름을 소스분할 시트에 출력하고, 결과가 성공적이면 작업결과에 SUCCESS라고 출력한다.

<img src="https://github.com/glowthem/ToyProject/blob/master/Git%20%EC%86%8C%EC%8A%A4%20%EC%97%85%EB%A1%9C%EB%93%9C%20%EC%9E%90%EB%8F%99%ED%99%94/imgSrc/%EC%86%8C%EC%8A%A4%EB%B6%84%ED%95%A0.png" width="90%" height="90%" alt="소스분할">
<br/>

> 2. *기업체이름_Solution(파일 번호)* 를 **이름** 으로, *프로그램 이름(파일 번호+파일 이름)* 을 **description** 으로 저장소 생성한다.<br/>
> HTTP POST 방식으로 json 전송. 결과를 GitLab 시트 내의 저장소 생성 결과에 출력한다. <br/>
> ❗️이미 존재하는 저장소를 생성하라는 request를 보낼 경우 결과로 Bad Request라고 출력된다.

<img src="https://github.com/glowthem/ToyProject/blob/master/Git%20%EC%86%8C%EC%8A%A4%20%EC%97%85%EB%A1%9C%EB%93%9C%20%EC%9E%90%EB%8F%99%ED%99%94/imgSrc/%EC%86%8C%EC%8A%A4%EB%B6%84%ED%95%A0.png" width="90%" height="90%" alt="GitLab결과"><br/>

> 3. 생성한 저장소를 미리 지정해놓은 경로 내에 *기업체이름_Solution(파일 번호)* 를 폴더명으로 하여 **클론 작업 수행** <br/>
> 4. **분리했던 소스파일들을 클론한 폴더로 이동** 하고 **저장소 내의 모든 파일 선택** (git add .) <br/>
> 5. 오늘 날짜를 **커밋 메시지** 로 하여 **커밋 수행** (git commit -m "2020-02-18") , **원격 저장소로 푸시** (git push -u origin master) <br/>
> 각 단계의 성공 여부를 Clone 결과, Commit 결과, Push 결과에 Success 또는 Failed로 출력