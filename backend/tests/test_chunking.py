import asyncio

from app.schemas import BookInfo, ChunkedProofreadRequest, ProofreadIssue, ProofreadLocator
from app.services import chunking


BOOK = BookInfo(title="测试书名", introduction="测试介绍")


def test_selection_under_threshold_uses_single_chunk():
    chunks = chunking.split_text_into_chunks("甲" * 7000, scope="selection")

    assert len(chunks) == 1
    assert chunks[0].index == 0
    assert chunks[0].start == 0
    assert chunks[0].end == 7000


def test_long_selection_splits_near_sentence_boundary():
    text = ("甲" * 4800) + "。" + ("乙" * 4800) + "。"

    chunks = chunking.split_text_into_chunks(text, scope="selection")

    assert len(chunks) == 2
    assert chunks[0].end == 4801
    assert chunks[0].text.endswith("。")
    assert chunks[1].start == chunks[0].end


def test_document_always_uses_chunking_and_prefers_paragraph_boundary():
    text = ("甲" * 4300) + "\n\n" + ("乙" * 4300)

    chunks = chunking.split_text_into_chunks(text, scope="document")

    assert len(chunks) == 2
    assert chunks[0].text.endswith("\n\n")
    assert chunks[1].start == chunks[0].end


def test_chunking_extends_to_next_sentence_boundary_instead_of_hard_cutting():
    text = ("甲" * 5000) + "。" + ("乙" * 1200) + "。"

    chunks = chunking.split_text_into_chunks(text, scope="document")

    assert len(chunks) == 2
    assert chunks[0].end == 5001
    assert chunks[0].text.endswith("。")
    assert chunks[1].start == chunks[0].end


def test_chunking_keeps_remaining_text_when_no_boundary_exists():
    text = "甲" * 8500

    chunks = chunking.split_text_into_chunks(text, scope="document")

    assert len(chunks) == 1
    assert chunks[0].start == 0
    assert chunks[0].end == len(text)


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
        locator=ProofreadLocator(
            key="有错字。",
            key_start=2,
            key_end=6,
            original_start_in_key=1,
            original_end_in_key=3,
            strategy="context",
            key_occurrence_index=0,
        ),
    )

    global_issue = chunking.globalize_issues(chunk, [issue])[0]

    assert global_issue.chunk_index == 2
    assert global_issue.start == 3
    assert global_issue.end == 5
    assert global_issue.global_start == 3003
    assert global_issue.global_end == 3005
    assert global_issue.locator is not None
    assert global_issue.locator.key_start == 3002
    assert global_issue.locator.key_end == 3006
    assert global_issue.locator.original_start_in_key == 1
    assert global_issue.locator.key_occurrence_index == 0


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
        text=("甲" * 4999) + "。" + ("乙" * 4999) + "。",
        book=BOOK,
        scope="document",
    )

    issues, completed_chunks, failed_chunks = asyncio.run(
        chunking.proofread_chunks(request, fake_proofread_chunk)
    )

    assert completed_chunks == 2
    assert failed_chunks == 0
    assert [issue.chunk_index for issue in issues] == [0, 1]
    assert [issue.global_start for issue in issues] == [0, 5000]
