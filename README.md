# GitHub Pages 업로드용 폴더

이 폴더는 GitHub 저장소에 그대로 올리기 쉽게 정리한 배포용 묶음입니다.

## 사용 방법

1. 새 GitHub 저장소를 만듭니다.
2. 이 폴더 안의 `docs/` 폴더와 `README.md`를 저장소 루트에 올립니다.
3. GitHub 저장소 설정에서 `Pages`를 엽니다.
4. Source를 `Deploy from a branch`로 선택합니다.
5. Branch는 `main`, Folder는 `/docs`를 선택합니다.
6. 저장 후 몇 분 기다리면 사이트 주소가 생성됩니다.

## 배포 후 주소 예시

```text
https://사용자명.github.io/저장소명/
```

## 포함 파일

- `docs/index.html`
- `docs/app.mjs`
- `docs/engine.mjs`
- `docs/data.mjs`
- `docs/.nojekyll`
- `docs/README.md`
