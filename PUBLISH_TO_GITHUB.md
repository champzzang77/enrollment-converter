# GitHub 업로드 순서

## 1. 이 폴더로 이동

```bash
cd /Users/parkchamin/vibe/github_pages_bundle
```

## 2. git 사용자 정보가 없으면 1회 설정

```bash
git config user.name "내 이름"
git config user.email "내 이메일"
```

## 3. 첫 커밋

```bash
git add .
git commit -m "Initial GitHub Pages site"
```

## 4. GitHub 새 저장소 연결

```bash
git remote add origin https://github.com/사용자명/저장소명.git
git push -u origin main
```

## 5. GitHub Pages 켜기

GitHub 저장소에서 아래처럼 설정합니다.

1. `Settings`
2. `Pages`
3. `Deploy from a branch`
4. Branch: `main`
5. Folder: `/docs`

배포 주소 예시:

```text
https://사용자명.github.io/저장소명/
```

## 참고

- 사이트 파일은 `docs/` 폴더 안에 있습니다.
- 이 폴더 전체를 GitHub 저장소 루트에 올리면 됩니다.
