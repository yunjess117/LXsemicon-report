"""유튜브 스튜디오 내보내기 파서. xlsx와 csv 둘 다 받는다.

실제 6월 샘플로 확인된 형식: '표 데이터' 시트
  컬럼: 콘텐츠(=video_id), 동영상 게시 시간, 동영상 제목, 조회수, 길이(초),
        평균 시청 지속 시간('H:MM:SS' 문자열), 평균 조회율 (%), 좋아요, 추가된 댓글 수, 공유
  첫 데이터 행은 A='합계' 인 요약 행이므로 건너뛴다.

같은 파일의 '차트 데이터' 시트에도 콘텐츠·동영상 제목 열이 있어서, 헤더 조건에 '조회수'를
넣어 '표 데이터' 쪽만 잡히게 한다(차트 데이터에는 조회수 열이 없다).

video_id 존재 자체가 '당월 발행분' 판별의 핵심 근거이므로, video_id 목록을
month_filter에서 그대로 재사용할 수 있게 반환한다.
"""

from dataclasses import dataclass

from .base import ParseError, find_column, find_table_with_header, to_number

HEADER_CANDIDATES = ["콘텐츠", "동영상 제목", "조회수"]

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
    header, data_rows = find_table_with_header(file, HEADER_CANDIDATES)
    if header is None:
        raise ParseError(
            "유튜브 스튜디오 원본에서 헤더 행(콘텐츠, 동영상 제목, 조회수 포함)을 찾지 못했습니다. "
            "스튜디오의 '표 데이터'를 내보낸 파일(xlsx 또는 csv)이 맞는지 확인해주세요."
        )

    col_idx = {}
    for key, candidates in COLS.items():
        idx = find_column(header, candidates)
        if idx is None:
            raise ParseError(f"유튜브 스튜디오 원본에서 '{candidates[0]}' 열을 찾지 못했습니다.")
        col_idx[key] = idx

    def cell(row, key):
        i = col_idx[key]
        return row[i] if i < len(row) else None

    out: list[YoutubeStudioRow] = []
    for row in data_rows:
        vid = cell(row, "video_id")
        if not vid or str(vid).strip() == "합계":
            continue
        duration = cell(row, "avg_view_duration")
        out.append(
            YoutubeStudioRow(
                video_id=str(vid).strip(),
                title=str(cell(row, "title") or "").strip(),
                views=to_number(cell(row, "views")),
                length_sec=to_number(cell(row, "length_sec")),
                avg_view_duration=str(duration).strip() if duration not in (None, "") else None,
                avg_view_pct=to_number(cell(row, "avg_view_pct")),
                likes=to_number(cell(row, "likes")),
                comments=to_number(cell(row, "comments")),
                shares=to_number(cell(row, "shares")),
            )
        )
    return out
