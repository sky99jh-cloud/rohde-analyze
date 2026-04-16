송신기의 log 데이터(html파일)를 파싱해서 원하는 데이터값을 추출후 엑셀로 저장하는 프로젝트

1. python 기반으로 작성할 것
2. gui 형식으로 html 파일과, excel 파일을 선택해서 작업할 것
3. excel 파일의 마지막 시트를 복사해서 끝에 붙여넣고 내용만 수정할 것
4. 추출값은 excel 파일의 마지막 시트에 추가하고, 다음에 excel 파일을 열었을때 마지막 시트가 기본으로 보여지도록 저장
5. excel 시트 내용중에 AMP 1 열에 들어갈 내용은 html 파일의 Output Stage » Rack 1 Amplifiers » Amplifier 1 » Supply, Transistors, RF Levels 참고해서 넣으면 됨. 
단, AMP Temp 항목은 Output Stage » Rack 1 Amplifiers » Amplifier 1 » Status 의 Amplifier Temp. 항목 참조해서 넣으면 됨
6. excel 시트 내용의 아랫부분 "특이사항" 에 Shoulder Distance 등의 값은 Exciter » Pre- Correction » Non Linear 와 Exciter » Pre- Correction » Linear 안의 해당되는 값을 참고해서 넣으면 됨.
7. excel 시트의 f3셀값과 i3셀값은  Power Limits » Power and Limits 항목의 Forward Power와 Reflected Power 값을 넣으면 됨
