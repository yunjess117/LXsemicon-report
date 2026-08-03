"""유튜브 스튜디오 내보내기(xlsx) 파서.

실제 6월 샘플로 확인된 형식: '표 데이터' 시트
  컬럼: 콘텐츠(=video_id), 동영상 게시 시간, 동영상 제목, 조회수, 길이(초),
        평균 시청 지속 시간('H:MM:SS' 문자열), 평균 조회율 (%), 좋아요, 추가된 댓글 수, 공유
  첫 데이터 행은 A='합계' 인 요약 행이므로 건너뛴다.

video_id 존재 자체가 '당월 발행분' 판별의 핵심 근거이므로, video_id 목록을
month_filter에서 그대로 재사용할 수 있게 반환한다.
"""

from dataclasses import dataclass

from .base import ParseError, find_column, find_sheet_by_pattern, load_sheet_rows

HEADER_CANDIDATES = ["콘텐츠", "동영상 제목"]

COLS = {
    "video_id": ["콘텐츠"],
    "title": ["동영상 제목"],
    "views": ["조회수"],
    "length_sec": ["길이"],
    "avg_view_duration": ["평균 시청 지속 시간"],
    "avg_view_pct": ["평균 조회율 (%)", "평균 조회율(%)"],
    "likes": ["좋아요"],
    "comments": ["추가된 댓글 수", "댓글"],
    "shares": ["공유"],
}


@dataclass
class YoutubeStudioRow:
    video_id: str
    title: str
    views: float | None
    length_sec: float | None
    avg_view_duration: str | None  # 'H:MM:SS' 원문 그대로
    avg_view_pct: float | None
    likes: float | None
    comments: float | None
    shares: float | None


def parse_youtube_studio(file) -> list[YoutubeStudioRow]:
    sheet_name = find_sheet_by_pattern(file, r"표\s*데이터")
    file.seek(0)
    _, rows = load_sheet_rows(file, sheet_name)

    header_idx = None
    for i, row in enumerate(rows[:5]):
        texts = {str(c).strip() for c in row if c is not None}
        if set(HEADER_CANDIDATES) <= texts:
            header_idx = i
            break
    if header_idx is None:
        raise ParseError("유튜브 스튜디오 '표 데이터' 시트에서 헤더 행을 찾지 못했습니다.")
    header = rows[header_idx]

    col_idx = {}
    for key, candidates in COLS.items():
        idx = find_column(header, candidates)
        if idx is None:
            raise ParseError(f"유튜브 스튜디오 원본에서 '{candidates[0]}' 열을 찾지 못했습니다.")
        col_idx[key] = idx

    out: list[YoutubeStudioRow] = []
    for row in rows[header_idx + 1 :]:
        vid = row[col_idx["video_id"]]
        if not vid or str(vid).strip() == "합계":
            continue
        out.append(
            YoutubeStudioRow(
                video_id=str(vid).strip(),
                title=str(row[col_idx["title"]] or "").strip(),
                views=row[col_idx["views"]],
                length_sec=row[col_idx["length_sec"]],
                avg_view_duration=row[col_idx["avg_view_duration"]],
                avg_view_pct=row[col_idx["avg_view_pct"]],
                likes=row[col_idx["likes"]],
                comments=row[col_idx["comments"]],
                shares=row[col_idx["shares"]],
            )
        )
    return out
