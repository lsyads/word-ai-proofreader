import asyncio

from app.schemas import BookInfo, ChunkedProofreadRequest, ProofreadIssue
from app.services import chunking


BOOK = BookInfo(title="测试书名", introduction="测试介绍")


def test_selection_under_threshold_uses_single_chunk():
    chunks = chunking.split_text_into_chunks("短文本。" * 10, scope="selection", chunk_size=3000)

    assert len(chunks) == 1
    assert chunks[0].index == 0
    assert chunks[0].start == 0
    assert chunks[0].end == len("短文本。" * 10)


def test_long_selection_splits_near_sentence_boundary():
    text = ("甲" * 2800) + "。" + ("乙" * 2800) + "。"

    chunks = chunking.split_text_into_chunks(text, scope="selection", chunk_size=3000)

    assert len(chunks) == 2
    assert chunks[0].end == 2801
    assert chunks[0].text.endswith("。")
    assert chunks[1].start == chunks[0].end


def test_document_always_uses_chunking_and_prefers_paragraph_boundary():
    text = ("甲" * 1800) + "\n\n" + ("乙" * 1800)

    chunks = chunking.split_text_into_chunks(text, scope="document", chunk_size=3000)

    assert len(chunks) == 2
    assert chunks[0].text.endswith("\n\n")
    assert chunks[1].start == chunks[0].end


def test_chunking_hard_cuts_when_no_boundary_exists():
    text = "甲" * 6500

    chunks = chunking.split_text_into_chunks(text, scope="document", chunk_size=3000)

    assert [chunk.start for chunk in chunks] == [0, 3000, 6000]
    assert [chunk.end for chunk in chunks] == [3000, 6000, 6500]


def test_globalize_issues_adds_chunk_offsets():
    chunk = chunking.ProofreadChunk(index=2, start=3000, end=6000, text="这里有错字。")
    issue = ProofreadIssue(
        id="issue-1",
        category="typo",
        severity="low",
        original="错字",
        replacement="改字",
        suggestion="修正错别字。",
        start=3,
        end=5,
    )

    global_issue = chunking.globalize_issues(chunk, [issue])[0]

    assert global_issue.chunk_index == 2
    assert global_issue.start == 3
    assert global_issue.end == 5
    assert global_issue.global_start == 3003
    assert global_issue.global_end == 3005


def test_proofread_chunked_aggregates_issues():
    async def fake_proofread_chunk(text, book, session_id, provider_api, proofread_mode):
        return [
            ProofreadIssue(
                id=f"issue-{text[0]}",
                category="style",
                severity="medium",
                original=text[:2],
                suggestion="建议",
                start=0,
                end=2,
            )
        ]

    request = ChunkedProofreadRequest(
        text=("甲" * 3000) + ("乙" * 3000),
        book=BOOK,
        scope="document",
        chunk_size=3000,
    )

    issues, completed_chunks, failed_chunks = asyncio.run(
        chunking.proofread_chunks(request, fake_proofread_chunk)
    )

    assert completed_chunks == 2
    assert failed_chunks == 0
    assert [issue.chunk_index for issue in issues] == [0, 1]
    assert [issue.global_start for issue in issues] == [0, 3000]
