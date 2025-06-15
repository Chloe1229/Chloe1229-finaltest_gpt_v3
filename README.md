선언문

[전체적인 작업목적]
이 작업명세서는 step8의 코딩 재구성과 수정에 대해 네 가지 카테고리로 구분하여 상세히 지시합니다.
각 카테고리별로 작업 세부지시 사항이 명확하게 구분되어 있으나, 모든 카테고리의 내용을 별개로 보지 말고 반드시 **종합적으로** 고려하여 step8의 최종 결과물에 **모두 반영**되어야 합니다.

[4가지 카테고리 작업지시에서 반드시 준수할 핵심원칙]
- step8의 표 자동화·생성 로직은 반드시 저장소에 있는 '제조방법변경 신청양식_empty_.docx' 파일을 기준 공양식으로 하며,
- step6, step7에서 사용된 **키명(key name)**, **저장경로(key path)**, **데이터 구조(dict/list)**를 반드시 그대로 활용하여  
  - 예시: step7_results[title_key][result_index]['output_1_tag'], step7_results[title_key][result_index]['output_2_text'], step6_items[title_key]['requirements'], step6_selections[f"{title_key}_req_{k}"] 등
- step8은 위 키 및 경로를 통해 확보된 모든 데이터를, 각 작업지시서별 세부 자동화 규칙에 따라  
  “제조방법변경 신청양식_empty_.docx”의 구조에 1:1로 자동 채워넣고(자동화 입력),  
  streamlit 화면에서는 이 표 구조가 **텍스트 기반 표 구조로 실시간 생성**되어야 하며,  
  사용자가 “파일 다운로드” 버튼을 클릭하면 동일 구조로 워드(.docx) 파일이 즉시 생성·다운로드 가능해야 합니다.
- streamlit 화면에 나타나는 표, 내용, 구조, 데이터와  
  다운로드된 워드파일 내 표, 내용, 구조, 데이터는  
  **완전히 동일**해야 하며,  
  “인쇄하기”를 누를 경우에는 streamlit이 사용자의 실제 윈도우/OS 프린터 환경과 직접 연동되어,  
  **현재 화면과 동일한 표/출력 결과**가 바로 인쇄될 수 있어야 합니다.
- 모든 페이징(paging) 및 표 렌더링(table rendering)은  
  **각 title_key에서 도출된 (output_1_tag, output_2_text) 조합별로  
  개별적으로 페이지·표가 생성**되어야 하며,
  - 즉, step7_results[title_key] 리스트의 각 인덱스(result_index)가  
    고유 조합(결과) 하나를 의미하며,  
    **(title_key, result_index)별로 반드시 한 페이지/한 표가 독립적으로 생성**되어야 합니다.
- 위 자동화 매핑 기준,  
  step8 구현의 모든 코드/키/경로/결과는  
  (title_key, result_index) 단위로 분기·반복되어,  
  결과가 누락, 병합, 생략, 중복 없이  
  명세된 대로 **개별 표/개별 페이지**로 구현되어야 합니다.

**이 규칙을 모든 step8 코드, 화면, 워드 자동화에 일관되게 적용해야 하며, 어느 한 항목도 누락/병합/생략/무시될 수 없습니다.**
---------
Declaration

[General Purpose of the Work]
This specification divides the detailed coding requirements for the restructuring and correction of step8 into four categories.
Although each category presents detailed instructions and explanations, all content must be applied comprehensively to the final implementation of step8, and should never be considered in isolation.

[Critical Principle Required Across All 4 Instruction Categories]
- The entire automation and table generation logic for step8 must use the '제조방법변경 신청양식_empty_.docx' file in the repository as the exact base template.
- All key names, key paths, and data structures (dict/list) from step6 and step7 must be directly referenced and mapped,  
  - Example: step7_results[title_key][result_index]['output_1_tag'], step7_results[title_key][result_index]['output_2_text'], step6_items[title_key]['requirements'], step6_selections[f"{title_key}_req_{k}"], etc.
- All data acquired from the above keys/paths must be auto-inserted into the template structure, in strict accordance with each category’s automation rules,  
  such that the [제조방법변경 신청양식_empty_.docx] template is filled out 1:1 (cell by cell, field by field) automatically.
- On the streamlit interface, this template table must be generated in real-time as a text-based table.  
  When the user clicks “파일 다운로드,” the system must create and provide a Word (.docx) file that is structurally and content-wise identical to the on-screen table.
- The table, content, and data shown on the streamlit page,  
  and the downloaded Word file,  
  **must be perfectly identical in structure, data, layout, and values.**  
  When “인쇄하기” is pressed, the print environment of streamlit must be linked to the user’s local printer/OS print dialog so that the output is exactly as shown on the screen.
- All paging and table rendering must be performed  
  **for each unique combination of output_1_tag and output_2_text for a given title_key**;  
  that is, each result (row) in the step7_results[title_key] list (with its own result_index)  
  represents a unique combination,  
  and **each (title_key, result_index) pair must generate a distinct page and table** independently.
- Following this automation mapping,  
  all code, keys, data routing, and rendered output in step8 must  
  branch and repeat strictly by (title_key, result_index),  
  with no omissions, no merging, no skipping, and no duplication,  
  so that every page and table matches the specification exactly as an individual, unique output.

**This rule must be consistently applied to all step8 code, UI, and Word automation, with no item ever omitted, merged, skipped, or ignored.**

----


# 첫번째 재구성 및 수정 요청사항

## 1. 목적 및 배경
- step8 페이지에서의 페이징 및 표 생성은 반드시 첨부된 워드 신청서 양식(공식 문서) 구조를 따라야 함.
- 기존에는 title_key 단위로 한 페이지만 생성되었으나,
- 실제 Step7 하드코딩 결과(output_1_tag, output_2_text)가 동일한 title_key에서 복수(2개 이상) 생성될 수 있음.
- 즉, 같은 title_key에서 output_1_tag/output_2_text 조합이 다르면, 해당 조합마다 별도의 페이지와 표를 생성해야 함.

---

## 2. 데이터 구조 요구

### [A] step7_results 구조 표준화
- 기존:
    step7_results = {
        "s2_2": {
            "title_text": "...",
            "output_1_tag": "...",
            "output_2_text": "..."
        },
        # ... (title_key별 단일 결과)
    }
- 변경:
    - title_key별로 복수 결과가 리스트로 존재해야 함
    - 예시:
        step7_results = {
            "s2_2": [
                {
                    "title_text": "...",
                    "output_1_tag": "...",
                    "output_2_text": "..."
                },
                {
                    "title_text": "...",
                    "output_1_tag": "...",
                    "output_2_text": "..."
                }
            ],
            # ...
        }
- 반드시 모든 title_key 값이 리스트(list)로 존재해야 하며, 결과가 1개여도 리스트로 래핑함.

---

## 3. 페이징/페이지 순회 구조

- step8의 페이지 리스트는 title_key, result_index 쌍의 일련번호로 구성해야 함.
- 예시:
    page_list = []
    for title_key, results in step7_results.items():
        for idx in range(len(results)):
            page_list.append((title_key, idx))
    # page_list 예시: [("s2_2", 0), ("s2_2", 1), ("p3_13", 0), ...]

---

## 4. 화면 및 워드 파일 표 생성 로직

- 각 페이지는 page_list[현재 인덱스]의 (title_key, idx) 쌍을 참조하여
    - step7_results[title_key][idx]에 있는 output_1_tag, output_2_text 값으로 표를 생성
    - 표 구조 및 제목, 항목 등은 첨부된 워드양식과 일치해야 하며,
    - 나머지 신청인, 충족조건, 필요서류 등 모든 데이터 매핑은 기존 로직과 동일하게 적용

---

## 5. Streamlit 및 워드 파일 동기화

- 화면에서 표시되는 표와 다운로드되는 워드 파일은 반드시 동일한 페이지/결과를 사용
- 다운로드 버튼 및 인쇄 버튼 등은 기존대로 페이지 단위로 연동

---

## 6. 예외처리 및 기타
- 복수 결과가 있을 때만 여러 페이지로 분기, 단일 결과도 항상 리스트 기반으로 접근
- 모든 반복문, 인덱싱, 표 생성 함수에서 title_key, idx 쌍을 기준으로 호출

---

## 7. 필수 코드 예시

### (1) step7_results 구조 처리

for k, v in step7_results.items():
    if not isinstance(v, list):
        step7_results[k] = [v]  # 결과 1개도 리스트화

page_list = []
for title_key, results in step7_results.items():
    for idx in range(len(results)):
        page_list.append((title_key, idx))

### (2) 현재 페이지별 데이터 접근

current_title_key, current_result_idx = page_list[current_page]
result = step7_results[current_title_key][current_result_idx]
# result["output_1_tag"], result["output_2_text"] 등 사용

### (3) 표 생성 함수 (워드/화면 공통)

def render_result_table(title_key, result):
    # result는 step7_results[title_key][idx]의 dict
    # 표 생성 로직 (신청인, 변경유형, 신청유형, 충족조건, 필요서류 등)
    ...

---

## 8. 수정 대상/위치

- step8에서 페이지 처리하는 모든 부분(페이지 리스트, 네비게이션, 데이터 바인딩, 표 생성 등)
- 워드 파일 생성 함수(create_application_docx 등)도 동일 로직 반영

---

## 9. 결과 확인

- 동일 title_key에서 output_1_tag/output_2_text가 여러 개 나오면, 각 조합별로 완전히 별도 페이지/표/워드파일이 생성되는지 반드시 검증

---

## 10. 요약

- title_key별 1페이지 → title_key+output조합별 1페이지로 구조를 변경
- 모든 반복, 인덱싱, 표, 워드생성은 (title_key, idx) 기반으로만 작동
- 데이터 접근, 표생성, 다운로드 등 모든 부분에서 이 구조를 일관 적용
- 코드 내 모든 참조부/반복부는 반드시 이 구조 반영

---

상기 사항 미준수 시 결과 인정 불가.
(불분명할 경우 본 작업지시서 예시/구조/용어만 엄격 준수)



----
# 첫번째 재구성 및 수정 요청사항(Eng. ver) First Restructuring and Revision Request

## 1. Purpose and Background
- The paging and table generation in the step8 page must strictly follow the structure of the attached official Word application form template.
- Previously, only one page was generated per title_key.
- However, in actual Step7 hardcoded results (output_1_tag, output_2_text), **multiple results (two or more) can be generated for the same title_key**.
- **In other words, even if the title_key is the same, if the output_1_tag/output_2_text combination differs, a separate page and table must be generated for each combination.**

---

## 2. Data Structure Requirements

### [A] Standardization of step7_results Structure
- **Previous structure:**  
    ```python
    step7_results = {
        "s2_2": {
            "title_text": "...",
            "output_1_tag": "...",
            "output_2_text": "..."
        },
        # ... (single result per title_key)
    }
    ```
- **Revised structure:**  
    - Each title_key must have a **list of results** (even if there is only one result)
    - Example:
    ```python
    step7_results = {
        "s2_2": [
            {
                "title_text": "...",
                "output_1_tag": "...",
                "output_2_text": "..."
            },
            {
                "title_text": "...",
                "output_1_tag": "...",
                "output_2_text": "..."
            }
        ],
        # ...
    }
    ```
- **All title_key values must always be lists**. If there is only one result, it must still be wrapped in a list.

---

## 3. Paging / Page Iteration Structure

- The page list in step8 must be constructed as a **tuple of (title_key, result_index)** for each result.
- Example:
    ```python
    page_list = []
    for title_key, results in step7_results.items():
        for idx in range(len(results)):
            page_list.append((title_key, idx))
    # Example: [("s2_2", 0), ("s2_2", 1), ("p3_13", 0), ...]
    ```

---

## 4. Screen and Word File Table Generation Logic

- For each page, reference the (title_key, idx) pair from page_list[current_page_index]
    - Use step7_results[title_key][idx] for output_1_tag and output_2_text for table generation.
    - The table structure, headings, and fields must match the attached Word template as closely as possible.
    - Other fields such as applicant info, satisfaction conditions, and required documents are mapped as per existing logic.

---

## 5. Streamlit and Word File Synchronization

- The table displayed on screen and the Word file downloaded **must use exactly the same page/result data**.
- The download and print buttons must remain page-specific, exactly as currently implemented.

---

## 6. Exception Handling and Additional Details

- If a title_key has zero results, that title_key must be excluded from page generation.
- If there are multiple results for a title_key, generate a page for each result. If there is only one, still use the list-based approach.
- All loops, indexing, and table generation functions must use the (title_key, idx) tuple for data referencing.

---

## 7. Required Code Examples

### (1) Ensuring step7_results Structure

```python
# Make sure all title_key values are lists
for k, v in step7_results.items():
    if not isinstance(v, list):
        step7_results[k] = [v]  # Wrap single result in a list

# Build page_list
page_list = []
for title_key, results in step7_results.items():
    for idx in range(len(results)):
        page_list.append((title_key, idx))

### (2) Accessing Data for the Current Page
current_title_key, current_result_idx = page_list[current_page]
result = step7_results[current_title_key][current_result_idx]
# Use result["output_1_tag"], result["output_2_text"], etc.


### (3) Table Generation Function (for both Word and Screen)
def render_result_table(title_key, result):
    # result is step7_results[title_key][idx]
    # Build the table: Applicant, Change Type, Application Type, Satisfaction Conditions, Required Documents, etc.
    ...


## 8. Targets/Locations for Modification
All sections of step8 that handle page iteration (page list, navigation, data binding, table rendering, etc.)
The Word file generation function (such as create_application_docx) must also be updated to use the same logic.

## 9. Result Verification
If multiple output_1_tag/output_2_text combinations are generated for the same title_key, there must be a fully separate page, table, and Word file for each combination.
Test with multi-result and single-result cases to ensure all logic is robust.

## 10. Summary
Change from "one page per title_key" to "one page per output combination (title_key, idx)".
All loops, indexing, table generation, and Word creation must be based exclusively on (title_key, idx).
Data access, table rendering, download functions, etc. must all use this structure consistently.
All references, iterations, and function calls in the code must strictly follow this new pattern.


------

# 두번째 재구성 및 수정 요청사항

## 1. 목적

- 첫번째 재구성 및 수정 요청사항에서 요구한 기준에 따라, step7에서 해당 (title_key, 조건) 조합에 맞는 도출 결과가 없는 경우 step8의 동작을 엄격히 규정하기 위함.

---

## 2. 요구 동작

- 특정 (title_key, 조건)에 대해 step7에서 도출 결과가 없을 경우(output_1_tag와 output_2_text가 모두 없는 경우):
    - 해당 step8 페이지에서는 워드파일 다운로드, 인쇄하기 기능을 **제공하지 않는다**.
    - 이 페이지에서 **화면에 표시되는 내용은 아래 문구**(줄바꿈 포함) 노출되어야 한다.

        해당 변경사항에 대한 충족조건을 고려하였을 때,
        「의약품 허가 후 제조방법 변경관리 가이드라인」에서 제시하고 있는
        범위에 해당하지 않는 것으로 확인됩니다.
* Move the code that displays the page number and renders the “⬅ 이전/다음 ➡” buttons outside the `else:` block in `step1_to_8_step8_final_.py`.
* Ensure these elements are shown regardless of whether `current_idx` is `None` or a valid index.
---

## 3. 화면 출력 로직

- step8의 어느 페이지에서든 step7에서 현재 (title_key, 조건)에 맞는 결과가 없으면:
    - **절대 표시하지 말아야 할 것:**
        - 워드파일 다운로드 버튼
        - 인쇄 버튼
        - 어떠한 표/양식 필드도
    - **지정 문구** 출력해야 함(위치·줄바꿈 등 포함).
    - Move the code that displays the page number and renders the “⬅ 이전/다음 ➡” buttons outside the `else:` block in `step1_to_8_step8_final_.py`.
    - Ensure these elements are shown regardless of whether `current_idx` is `None` or a valid index.
---

## 4. 페이징/이동

- 이 “결과 없음” 페이지도 step8 페이징에 포함되어야 하며, 다른 페이지처럼 “이전/다음” 네비게이션은 정상적으로 동작해야 한다.

---

## 5. 코드 예시

# step8 출력 로직 예시(Python 유사 코드):
if not has_matching_step7_result:
    st.write(
        "해당 변경사항에 대한 충족조건을 고려하였을 때,\n"
        "「의약품 허가 후 제조방법 변경관리 가이드라인」에서 제시하고 있는\n"
        "범위에 해당하지 않는 것으로 확인됩니다"
    )
* Move the code that displays the page number and renders the “⬅ 이전/다음 ➡” buttons outside the `else:` block in `step1_to_8_step8_final_.py`.
* Ensure these elements are shown regardless of whether `current_idx` is `None` or a valid index.
    # 다운로드/인쇄/표는 절대 출력 금지
else:
    # 정상 표/버튼/출력 로직
    ...

---

## 6. 요약

- step7에 해당 (title_key, 조건) 조합 결과가 없는 경우
    - step8 화면에는 반드시 지정 문구 노출
    - 워드파일 다운로드 및 인쇄 기능은 비활성화/숨김 처리
    - 표/양식 표시하지 않음
- 이 로직은 “결과 없음” 페이지에 대해 예외 없이 엄격히 적용되어야 한다.
- Move the code that displays the page number and renders the “⬅ 이전/다음 ➡” buttons outside the `else:` block in `step1_to_8_step8_final_.py`.
- Ensure these elements are shown regardless of whether `current_idx` is `None` or a valid index.
---

상기 요구사항 미준수시 결과물은 인정되지 않음.  
예시/지침/메시지내용은 반드시 완전히 일치하게 적용할 것.

---

# 두번째 재구성 및 수정 요청사항(Eng. ver) Second Restructuring and Revision Request

## 1. Purpose

- To define strict behavior for step8 when no matching output result exists for a given (title_key, condition) combination in step7, according to the requirements of the first restructuring request.

---

## 2. Required Behavior

- When, for a given (title_key, condition), no output result is generated in step7 (i.e., no matching output_1_tag and output_2_text is present):
    - On the relevant step8 page, do not provide Word file download or print functions.
    - The only content to be displayed on the screen for this page must be the following message (with line breaks preserved):

    - (Korean original for reference:  
        "해당 변경사항에 대한 충족조건을 고려하였을 때,\n「의약품 허가 후 제조방법 변경관리 가이드라인」에서 제시하고 있는\n범위에 해당하지 않는 것으로 확인됩니다")
    - Move the code that displays the page number and renders the “⬅ 이전/다음 ➡” buttons outside the `else:` block in `step1_to_8_step8_final_.py`.
    - Ensure these elements are shown regardless of whether `current_idx` is `None` or a valid index.
---

## 3. Screen Display Logic

- For any step8 page where no result from step7 matches the current (title_key, condition):
    - Do NOT display:
        - Word file download button
        - Print button
        - Any table or form fields
    - DO display the specified message (see above).
    - Move the code that displays the page number and renders the “⬅ 이전/다음 ➡” buttons outside the `else:` block in `step1_to_8_step8_final_.py`.
    - Ensure these elements are shown regardless of whether `current_idx` is `None` or a valid index.

---

## 4. Paging/Navigation

- This "no result" page must still be included in the step8 paging sequence, just like any other page.
- Navigation ("Previous", "Next") must work as usual.

---

## 5. Code Example

# Pseudocode for step8 display logic:
if not has_matching_step7_result:
    st.write(
        "Considering the satisfaction conditions for this change,\n"
        "it has been confirmed that it does not fall within the scope\n"
        "presented in the \"Guidelines for Post-Approval Changes in Drug Manufacturing Methods.\""
    )
    # Do NOT render download/print buttons or tables
    # Move the code that displays the page number and renders the “⬅ 이전/다음 ➡” buttons outside the `else:` block in `step1_to_8_step8_final_.py`.
    # Ensure these elements are shown regardless of whether `current_idx` is `None` or a valid index.
else:
    # Normal rendering of table, download, print, etc.
    ...

---

## 6. Summary

- When there is no matching result in step7 for the current (title_key, condition),
    - step8 screen must  display the designated message,
    - Word file download and print options are disabled/hidden for that page,
    - no table or form is rendered.
    - Move the code that displays the page number and renders the “⬅ 이전/다음 ➡” buttons outside the `else:` block in `step1_to_8_step8_final_.py`.
    - Ensure these elements are shown regardless of whether `current_idx` is `None` or a valid index.
- This logic must be strictly and consistently applied for all cases where no step7 result is found for the page.

---

Failure to fully comply with these requirements will result in rejection of the output.
Follow the instructions, examples, and message content exactly as provided.

---

# 세번째 재구성 및 수정 요청사항

## 1. 목적

- step8 페이지의 상단 UI, 버튼, 제목, 페이지네이션, 충족조건·오류·표시 등 시각적, 기능적 배치와 동작을 첨부 그림과 상세 요구에 따라 완전히 수정하는 것.

---

## 2. 상세 요구 및 기능/레이아웃 개선

① **상단 파일다운로드 버튼**
   - 좌측 상단에 “파일 다운로드” 버튼 배치.
   - 버튼 텍스트가 두 줄로 줄바꿈되거나 잘리지 않게 반드시 한 줄로만 표시. (버튼 너비/텍스트 크기 조절)
   - 버튼 주변 여백은 과도하지 않게 적절히 조정.

② **상단 인쇄하기 버튼**
   - 우측 상단에 “인쇄하기” 버튼 배치.
   - 버튼 텍스트가 반드시 한 줄로만 표시되도록, 버튼 너비와 텍스트 크기 조절.
   - 버튼 주변 여백은 과도하지 않게 조정.
   - 인쇄하기 버튼 클릭 시 실제 사용자 환경에서 프린트 창이 열리도록 기능 정상화(window.print() 등 브라우저 표준 방식 완벽 연동, Streamlit 환경에서도 동작 보장).

③ **상단 제목·페이지 번호 배치**
   - “파일 다운로드”와 “인쇄하기” 버튼은 첫 번째 줄(최상단, 좌/우측)에만 위치. (중앙은 비워둠)
   - **두 번째 줄에 중앙 정렬로** “「의약품 허가 후 제조방법 변경관리 가이드라인(민원인 안내서)」[붙임] 신청양식 예시” 제목만 단독 한 줄 배치. 버튼 줄과 분리!
   - **세 번째 줄에는** 페이지 번호(예: “1 / 7”)만 중앙정렬로 제목 아래 별도 한 줄 배치.
   - 어떤 경우에도 버튼 줄에 제목·페이지 번호가 오지 않도록 할 것. (줄바꿈/정렬 필수)
   - Display the page number immediately beneath the title and above the table generation code, so that it forms the third header row.
   - Confirm that the buttons remain on the first row, the title on the second, and the table begins below the page number.

④ **오류 태그/불필요 HTML 출력 방지**
   - 각 페이지 상단에 <br><h5>1. 신청인</h5> 등 HTML 태그가 코드 그대로 노출되지 않게.
   - 표/텍스트 등 실제 데이터만 노출되고, HTML 코드는 모두 렌더링된 결과로만 보여야 함.
   - Streamlit에서 `st.markdown(..., unsafe_allow_html=True)` 등 사용 시 HTML 코드를 그대로 노출하는 버그 해결.

⑤ **충족조건 표 누락 오류 해결**
   - “충족조건” 표는 동일 title_key의 step6 {requirements}를 기반으로, 조건 개수만큼 줄 생성.
   - 각 줄에는 “충족조건” 열과 “조건 충족 여부” 열 포함.
   - 각 조건에 대해 step6_items[title_key]["requirements"][키]의 텍스트, step6_selections[title_key+"_req_"+키]가 “충족”이면 ○, “미충족”이면 ×로 출력.
   - 충족조건 표 내용이 누락 없이 전부 생성되어야 하며, 표의 구조(머릿줄, 각 줄, 정렬)는 첨부 워드 양식과 동일하게 구성.

⑥ **페이지 구조 요약**
   - 1줄: 좌 “파일 다운로드” 버튼 / 우 “인쇄하기” 버튼 (한 줄, 중간 비움)
   - 2줄: 제목 “「의약품 허가 후 제조방법 변경관리 가이드라인(민원인 안내서)」[붙임] 신청양식 예시” (중앙, 단독 한 줄)
   - 3줄: 페이지 번호(“현재/총페이지”) (중앙, 단독 한 줄)
   - 그 이후: 각 페이지별 표/항목 내용(불필요 HTML, 태그, 오류 문구 등 없음, 모든 표는 첨부 워드 양식 구조와 일치)
   - 버튼/제목/페이지번호가 한 줄에 섞여서 나오지 않게 반드시 분리.
   - Display the page number immediately beneath the title and above the table generation code, so that it forms the third header row.
   - Confirm that the buttons remain on the first row, the title on the second, and the table begins below the page number.
⑦ **레이아웃 아스키 예시**

┌──────────────────────────────────────────────┐
│ [파일 다운로드]                     [인쇄하기] │ ← (첫 번째 줄, 양쪽에 버튼. 제목 없음)
│         「의약품 허가 후 제조방법 변경관리 가이드라인(민원인 안내서)」[붙임] 신청양식 예시         │ ← (두 번째 줄, 중앙, 제목만)
│                        1 / 7                 │ ← (세 번째 줄, 중앙, 페이지 번호만)
└──────────────────────────────────────────────┘

---

## 3. 추가 안내

- 각 항목(버튼, 제목, 페이지번호, 인쇄, 표, 오류 등)은 동시에 적용되어야 함.
- 요청한 모든 사항이 100% 일치하게 구현되지 않을 경우 결과물 불인정.
- 각 요구사항의 충족여부를 세부적으로 검증할 예정이므로, 한 줄, 한 칸, 한 글자도 요구와 다르면 안 됨.
- UI 및 기능 구현 시 첨부된 그림(스크린샷)과 완전히 동일한 구조로 배치할 것. (단, 위의 조정안이 더 우선)

----
# 세번째 재구성 및 수정 요청사항(Eng. ver) Third Restructuring and Revision Request

## 1. Purpose

- To completely revise the step8 page UI, buttons, title, pagination, satisfaction requirements, error displays, and all visual/functional layouts and behaviors in accordance with the attached screenshots and the detailed requirements below.

---

## 2. Detailed Requirements and UI/Functional Improvements

① **File Download Button (Top-Left)**
   - Place a "[파일 다운로드]" button at the top left.
   - The button label must always appear on a single line (never split/wrapped). Adjust button width or font size as needed.
   - Ensure button margins are not excessive.

② **Print Button (Top-Right)**
   - Place a "[인쇄하기]" button at the top right, on the same horizontal line as the file download button.
   - The label must always display on a single line. Adjust button width or font size if necessary.
   - Margins around the button must not be excessive.
   - The print button must open the actual print dialog on the user's computer (implement using standard browser-compatible approaches such as `window.print()`; must function correctly in Streamlit regardless of environment).

③ **Page Title and Pagination Display**
   - **First Row:** Only the "[파일 다운로드]" (left) and "[인쇄하기]" (right) buttons, with the rest of the row empty.
   - **Second Row:** Centered, display only the page title:
     - "「의약품 허가 후 제조방법 변경관리 가이드라인(민원인 안내서)」[붙임] 신청양식 예시"
     - The title must always be on its own line, not sharing the row with any button or page number.
   - **Third Row:** Directly below the title, center-aligned, display the page number (e.g., "1 / 7") on its own line.
   - Under no circumstances should the buttons, title, or page number be displayed on the same line together. Each element must occupy its own distinct row.

④ **Prevent Raw HTML/Tag/Error Output**
   - Do NOT display any raw HTML tags or unrendered code at the top of each page (e.g., `<br><h5>1. 신청인</h5>...` etc).
   - Only properly rendered tables and text data should be visible to the user.
   - Move the code that displays the page number and renders the “⬅ 이전/다음 ➡” buttons outside the `else:` block in `step1_to_8_step8_final_.py`.
   - Ensure these elements are shown regardless of whether `current_idx` is `None` or a valid index.
   - When using Streamlit (e.g., `st.markdown(..., unsafe_allow_html=True)`), ensure that raw HTML is rendered, not printed as literal text.

⑤ **Satisfaction Requirements Table – Omission/Error Fix**
   - The "Satisfaction Requirements" table must:
     - For the current `title_key`, use all requirements from step6 (`step6_items[title_key]["requirements"]`).
     - Generate a row for each requirement, with two columns: "Requirement" and "Met or Not Met".
     - For each row:
       - Display the requirement text.
       - Display "○" if `step6_selections[title_key+"_req_"+key] == "충족"`, or "×" if "미충족".
   - No requirement row may be omitted.
   - Table structure (headers, rows, alignment) must strictly match the official Word form template.

⑥ **UI Layout Summary**
   - Row 1: Left – [파일 다운로드] button, Right – [인쇄하기] button. No title or page number here.
   - Row 2: Centered page title only (Korean as above), as a single line.
   - Row 3: Centered page number (e.g., "1 / 7"), as a single line.
   - Following rows: Page content, tables, etc. (No raw HTML, tags, or extraneous error messages.)
   - Buttons, title, and page number must never share a row – always separate.

⑦ **UI Layout Example (ASCII Art)**

┌──────────────────────────────────────────────┐
│ [파일 다운로드]                     [인쇄하기] │ ← (First row: both buttons on sides, no title)
│         「의약품 허가 후 제조방법 변경관리 가이드라인(민원인 안내서)」[붙임] 신청양식 예시         │ ← (Second row: centered title only)
│                        1 / 7                 │ ← (Third row: centered page number only)
└──────────────────────────────────────────────┘

---

## 3. Additional Notes

- All changes (buttons, title, page number, print, table rendering, error prevention) must be applied together.
- Output will be rejected if any single requirement here is not followed exactly.
- All details (line, cell, margin, header, table, etc.) will be strictly checked – any deviation will result in rejection.
- The UI and layout must match the attached screenshots precisely unless a more explicit textual instruction above overrides them.


----



# 네번째 재구성 및 수정 요청사항
## 1. 목적

- step8의 각 페이지 및 다운로드 워드파일이 반드시 [제조방법변경 신청양식_empty_.docx]와 구조·폰트·셀 병합·열 너비·행 구성·내용·순서·셀 데이터까지 완전히 동일하게 생성되어야 하며,
- 표/자동화/페이지/조합별 데이터 도출시 실제 코드에서 사용하는 key, 인덱스, 변수명을 모두 명확하게 표기해야 함.
- 화면/워드의 표/데이터/셀 구조가 공양식과 100% 동일하게 보이고, 실제 값도 완벽 자동화로 들어가야 함.

---

## 2. 전체 표 구조 및 각 항목별 실제 자동화 매핑 (키·코드 예시 포함)

────────────────────────────────────────────
│ [파일 다운로드]                     [인쇄하기] │ ← (상단, 두 버튼. 한글로 표기)
│         「의약품 허가 후 제조방법 변경관리 가이드라인(민원인 안내서)」[붙임] 신청양식 예시         │ ← (중앙 제목, 한글 그대로)
│                        1 / N                │ ← (페이지 번호, 자동)
────────────────────────────────────────────


─────────────────────────────────────────────────────
│ 1. 신청인 (세로 3행 병합, 12pt, Bold) │ "성명" (11pt)        │ "" (공란, 입력값 없음)   │
│                                       │ "제조소(영업소) 명칭" (11pt) │ "" (공란)              │
│                                       │ "변경신청 제품명" (11pt)     │ "" (공란)              │
─────────────────────────────────────────────────────

- 좌측 항목명(3행): "성명", "제조소(영업소) 명칭", "변경신청 제품명"
- 실제 데이터 자동화: 모두 (공란) 표시, 고정
- 표 구조: 3행 2열, 셀 병합/열 너비/폰트 크기는 공양식과 동일
- **코드상 자동화:** 직접 데이터 없음, 고정 출력(예: `st.table`/워드 표에 직접 문자열)

─────────────────────────────────────────────────────────
│ 2. 변경유형 (12pt)   │ 3. 신청 유형(AR, IR, Cmin, Cmaj 중 선택) (12pt) │
─────────────────────────────────────────────────────────
│ {step7_results[title_key][idx]['title_text']} (11pt)        │ {step7_results[title_key][idx]['output_1_tag']} (11pt)        │
─────────────────────────────────────────────────────────

- **좌측(2. 변경유형):**
    - 자동화 값:  
      `step7_results[title_key][idx]['title_text']`
        (예: "3.2.P.1 완제의약품의 성상 및 조성\n9. 완제의약품(고형제제 제외)의 조성 변경")
- **우측(3. 신청유형):**
    - 첫줄:  
      `step7_results[title_key][idx]['output_1_tag']`
        (예: "Cmaj", "AR" 등)

- **표 구조:** 2행 2열(실제는 한 행, 각 셀 줄바꿈 포함), 공양식 너비, 줄간격, 폰트(제목 12pt, 내용 11pt)

────────────────────────────────────────────────────────────
│ 4. 충족조건 (12pt)          │ 조건 충족 여부(○, X 중 선택) (12pt)      │
────────────────────────────────────────────────────────────
│ {step6_items[title_key]['requirements'][k]} (11pt) │ {if step6_selections[f"{title_key}_req_{k}"] == '충족':표시 = '○'if step6_selections[f"{title_key}_req_{k}"] == '미충족':표시 = '×'} (11pt) │
│ ... (requirement 개수만큼 반복, 자동확장) ...        │  ... (requirement 개수만큼{if step6_selections[f"{title_key}_req_{k}"] == '충족':표시 = '○'if step6_selections[f"{title_key}_req_{k}"] == '미충족':표시 = '×'} (11pt) 반복, 자동확장) ...│
────────────────────────────────────────────────────────────

- **좌측:** requirements 항목(예: "1. ~조건명"), 자동 for문으로 행 생성
    - 실제 코드:  
      ```python
      for k, v in step6_items[title_key]['requirements'].items():
          좌측 = v
          우측 = '○' if step6_selections[f"{title_key}_req_{k}"] == '충족' else '×'
      ```
- **표 구조:** n행 2열, requirements 개수만큼 행 생성(11pt), 머릿줄(12pt)

────────────────────────────────────────────────────────────
│ 5. 필요서류 (해당 필요서류 기재) (12pt)      │ 구비 여부 (○, X 중 선택) / 해당 페이지 표시 (12pt) │
────────────────────────────────────────────────────────────
│ {각 행: output_2_text에서 줄바꿈 기준 분할} (11pt), │ (항상 공란, 자동)                                           │
│ ... 필요서류 행수만큼 ...                           │ ...                                                         │
────────────────────────────────────────────────────────────
    - "필요서류는 다음과 같습니다.'\n" 부분은 첫번째 행에 도출되지 않게 해주세요.

- **좌측:**  
    - 현재 페이지 조합의  
      `step7_results[title_key][idx]['output_2_text']`  
    - 줄바꿈 `\n` 기준 분할, 각 줄 한 행씩(11pt)
- **우측:**  
    - 각 행 모두 공란(직접 기입 없음)
- **표 구조:** n행 2열, 필요서류 개수만큼 행 생성, 머릿줄(12pt)
    - "필요서류는 다음과 같습니다.'\n" 부분은 첫번째 행에 도출되지 않게 해주세요.

---

### 실제 코드 키/경로 기반 명세서 (공양식 100% 일치)

─────────────────────────────────────────────
│ [파일 다운로드]                     [인쇄하기] │ ← Streamlit/Docx 모두 동일, 한글로만
│ 「의약품 허가 후 제조방법 변경관리 가이드라인(민원인 안내서)」[붙임] 신청양식 예시 │
│                   {current_page+1} / {total_pages}                    │
─────────────────────────────────────────────

─────────────────────────────────────────────────────
│ 1. 신청인 (세로 3행 병합, 12pt, Bold) │ "성명" (11pt)        │ "" (공란, 입력값 없음)   │
│                                       │ "제조소(영업소) 명칭" (11pt) │ "" (공란)              │
│                                       │ "변경신청 제품명" (11pt)     │ "" (공란)              │
─────────────────────────────────────────────────────
# 실제 코드 예시:
applicant_table = [
    ["1. 신청인", "성명", ""],
    ["", "제조소(영업소) 명칭", ""],
    ["", "변경신청 제품명", ""]
]
# 열 너비/행 높이/폰트 – docx, Streamlit 표 렌더시 공양식과 일치

─────────────────────────────────────────────────────────────
│ 2. 변경유형 (12pt, Bold)                                   │ 3. 신청유형(AR, IR, Cmin, Cmaj 중 선택) (12pt, Bold)        │
│───────────────────────────────────────────────────────────│
│ step7_results[title_key][idx]['title_text']   (11pt)      │ step7_results[title_key][idx]['output_1_tag'] (11pt)       │
─────────────────────────────────────────────────────────────

# 실제 코드:
# 좌측: step7_results[title_key][idx]['title_text']
# 우측 1행: step7_results[title_key][idx]['output_1_tag']

─────────────────────────────────────────────────────────────
│ 4. 충족조건 (12pt, Bold)                                         │ 조건 충족 여부 (○, X, 12pt, Bold)        │
│─────────────────────────────────────────────────────────────────│
│ 반복 for k, v in step6_items[title_key]['requirements'].items(): │
│   v (11pt)                                                      │ {if step6_selections[f"{title_key}_req_{k}"] == '충족':표시 = '○' if step6_selections[f"{title_key}_req_{k}"] == '미충족':표시 = '×'} (11pt) │
│ ... (requirements 개수만큼 반복, 행 자동확장)                     │ ... (requirement 개수만큼{if step6_selections[f"{title_key}_req_{k}"] == '충족':표시 = '○'if step6_selections[f"{title_key}_req_{k}"] == '미충족':표시 = '×'} (11pt) 반복, 자동확장) ...│
─────────────────────────────────────────────────────────────

# 실제 코드:
# for k, v in step6_items[title_key]['requirements'].items():
#    좌측 = v
#    우측 = '○'if step6_selections[f"{title_key}_req_{k}"] == '충족 and '×' if step6_selections[f"{title_key}_req_{k}"] == '미충족'

─────────────────────────────────────────────────────────────
│ 5. 필요서류 (12pt, Bold)      │ 구비 여부 (○, X, 12pt, Bold) / 해당 페이지 표시    │
│───────────────────────────────────────────────────────────│
│ for line in step7_results[title_key][idx]['output_2_text'].split('\n'): │
│   line (11pt)                                            │   "" (공란, 자동입력 없음)                    │
│ ... (필요서류 개수만큼 반복, 행 자동확장)                 │
─────────────────────────────────────────────────────────────
 **"필요서류는 다음과 같습니다.'\n" 부분은 첫번째 행에 도출되지 않게 해주세요. **
# 실제 코드:
# for line in step7_results[title_key][idx]['output_2_text'].split('\n'):
#     좌측 = line
#     우측 = ""

─────────────────────────────────────────────────────────────
# 전체 페이지별, 조합별(75개 케이스) 마다 아래와 같은 매핑
# - title_key, idx: page_list[페이지 인덱스]에서 결정
# - step7_results[title_key][idx]: 각 조건/조합에 대한 결과
# - step6_items[title_key]['requirements']: requirements 목록
# - step6_selections[f"{title_key}_req_{k}"]: 충족/미충족 상태
# 표 스타일, 병합, 열 너비, 폰트 크기/스타일, 정렬, 여백, 줄간격은  
# 반드시 제조방법변경 신청양식_empty_.docx의 스타일, 구조와 동일하게 구현

─────────────────────────────────────────────────────────────
# 예시 코드 (페이지 단위 자동화 표 렌더링 – Streamlit/Python)

def render_applicant_table():
    return [
        ["1. 신청인", "성명", ""],
        ["", "제조소(영업소) 명칭", ""],
        ["", "변경신청 제품명", ""]
    ]

def render_change_type_table(step7_results, title_key, idx):
    return [
        ["2. 변경유형", "3. 신청유형(AR, IR, Cmin, Cmaj 중 선택)"],
        [
            step7_results[title_key][idx]['title_text'],
            f"{step7_results[title_key][idx]['output_1_tag']}}"
        ]
    ]

def render_requirements_table(step6_items, step6_selections, title_key):
    rows = []
    for k, v in step6_items[title_key]['requirements'].items():
        mark = '○'if step6_selections[f"{title_key}_req_{k}"] == '충족 and '×' if step6_selections[f"{title_key}_req_{k}"] == '미충족'
        rows.append([v, mark])
    return [["4. 충족조건", "조건 충족 여부"]] + rows

def render_documents_table(step7_results, title_key, idx):
    lines = step7_results[title_key][idx]['output_2_text'].split('\n')
    rows = [[line, ""] for line in lines if line.strip()]
    return [["5. 필요서류 (해당하는 필요서류 기재)", "구비 여부 (○, X 중 선택) / 해당 페이지 표시"]] + rows
 **"필요서류는 다음과 같습니다.'\n" 부분은 첫번째 행에 도출되지 않게 해주세요. **


─────────────────────────────────────────────────────────────

# 모든 표/셀/값은 위 함수의 변수 및 키로, 실제 코드상 접근(임의 설명 불허)
# 워드파일(.docx), Streamlit 화면 모두 구조/내용/스타일 100% 일치 필수

─────────────────────────────────────────────────────────────

# 요약 및 주의
- 자연어 설명·추상표현 X
- 표/셀/내용/자동화 키·경로·구조·코드 예시로만 기술
- 공양식 기준 미세구조(폰트, 병합, 여백, 줄간격 등) 구현 누락 불가
- (title_key, idx)별로 반드시 개별 표 구성/데이터 자동화

---

## 3. 전체 코드/데이터 접근법 요약 (실제 구현 참고)

- title_key, idx:  
    - page_list = [(title_key, idx), ...]  
    - 현재 페이지에서 title_key, idx 참조
- step7_results:  
    - {title_key: [ { 'title_text', 'output_1_tag', 'output_2_text' }, ... ] }
- step6_items:  
    - {title_key: { 'requirements': {k: v, ... } } }
- step6_selections:  
    - {f"{title_key}_req_{k}": '충족'/'미충족', ... }

---

## 4. 예시(실제 한 페이지 자동화 데이터 바인딩)

예시: title_key="p1_9", idx=0

- 2. 변경유형: step7_results['p1_9'][0]['title_text']
- 3. 신청유형(AR/IR/Cmin/Cmaj): step7_results['p1_9'][0]['output_1_tag']
- 4. 충족조건:
    - for k, v in step6_items['p1_9']['requirements'].items():
        - 좌: v
        - 우: '○'if step6_selections[f"{title_key}_req_{k}"] == '충족 and '×' if step6_selections[f"{title_key}_req_{k}"] == '미충족'
- 5. 필요서류:
    - for line in step7_results['p1_9'][0]['output_2_text'].split('\n'):
        - 좌: line
        - 우: 공란
    - "필요서류는 다음과 같습니다.'\n" 부분은 첫번째 행에 도출되지 않게 해주세요.

---

## 5. 개발·테스트·문서화 지침

- 표/셀/머릿줄/행/열 너비/폰트/병합/간격: [제조방법변경 신청양식_empty_.docx] 구조와 100% 일치, 화면/워드 모두
- 자동화 매핑·코드·키명: 위 표 및 코드, 변수, 구조에 근거, 반드시 주석/명시
- 조합별 매칭: (title_key, idx)별로 위 구조 반복 생성
- 누락/오류/불일치 발생시 전체 작업 불인정

---

# 반드시 위 항목 전체가 실제 코드, 표, 데이터 매핑에 적용되어야 하며  
한 항목도 누락 없이 개발, 테스트, 코드리뷰, 산출물 문서화에 활용할 것.



정리하면,
위 모든 자동화 매핑, 코드 경로, 셀 구조, 항목별 데이터 입력 방식이
아래 예시처럼 “단 하나의 표(table)” 내에서
세로로 모든 항목이 연결된 상태로 완성되어야 한다.

─────────────────────────────────────────────────────────────
│ 1. 신청인 (세로 3행 병합, 12pt, Bold) │ 성명 (11pt) │ "" (공란) │
│ │ 제조소(영업소) 명칭 │ "" (공란) │
│ │ 변경신청 제품명 │ "" (공란) │
├───────────────────────────────────────┼───────────────────────┼──────────────────┤
│ 2. 변경유형 (세로 2행 병합, 12pt, Bold) │ 3. 신청 유형(AR, IR, Cmin, Cmaj 중 선택) (12pt, Bold) │
│ step7_results[title_key][idx]['title_text'] (11pt) │ step7_results[title_key][idx]['output_1_tag'] (11pt) │
├───────────────────────────────────────┴───────────────────────┴──────────────────┤
│ 4. 충족조건 (머릿셀, 12pt, Bold, 행 병합) │ "충족조건" (11pt, Bold) │ "조건 충족 여부(○, X 중 선택)" (11pt, Bold) │
│ 반복 for k, v in step6_items[title_key]['requirements'].items(): │ v (11pt) │'○'if step6_selections[f"{title_key}_req_{k}"] == '충족 and '×' if step6_selections[f"{title_key}_req_{k}"] == '미충족' (11pt) │
├─────────────────────────────────────────────────────────────┼──────────────────┤
│ 5. 필요서류 (머릿셀, 12pt, Bold, 행 병합) │ "필요서류 (해당하는 필요서류 기재)" (11pt, Bold) │ "구비 여부 (○, X 중 선택)" (11pt, Bold) │
│ for line in step7_results[title_key][idx]['output_2_text'].split('\n'): │ line (11pt) │ "" (공란) │
─────────────────────────────────────────────────────────────
** "필요서류는 다음과 같습니다.'\n" 부분은 첫번째 행에 도출되지 않게 해주세요.**

이렇게 “하나의 표” 내에서

모든 셀 병합, 폰트, 정렬, 열 너비/행 높이,

각 자동화 항목의 입력 방식(코드 경로/키)

충족조건·필요서류의 행 확장(동적) : 다만, "필요서류는 다음과 같습니다.'\n" 부분은 첫번째 행에 도출되지 않게 해주세요.


[제조방법변경 신청양식_empty_.docx] 구조와 100% 동일
이 반드시 적용되어야 하며,
“여러 개 표 분리, 제목+표 따로 생성”은 절대 불가함.
----


# 네번째 재구성 및 수정 요청사항(Eng. ver) Fourth Revision and Correction Instruction

## 1. Purpose

- Each page in step8 and the downloaded Word file must be generated to be **exactly identical** to [제조방법변경 신청양식_empty_.docx] in structure, font, cell merging, column width, row configuration, content, order, and cell data.
- When extracting data per table/automation/page/combination, all keys, indices, and variable names actually used in the code must be clearly stated.
- Both the screen and Word table/data/cell structure must be 100% identical to the template, and every value must be filled automatically with no discrepancy.

---

## 2. Overall Table Structure and Actual Automation Mapping for Each Item (Including Key/Code Examples)

────────────────────────────────────────────
│ [파일 다운로드]                     [인쇄하기] │ ← (Top, both buttons. In Korean only)
│         「의약품 허가 후 제조방법 변경관리 가이드라인(민원인 안내서)」[붙임] 신청양식 예시         │ ← (Center title, keep as Korean)
│                        1 / N                │ ← (Page number, auto)
────────────────────────────────────────────

─────────────────────────────────────────────────────
│ 1. 신청인 (vertical merge of 3 rows, 12pt, Bold) │ "성명" (11pt)        │ "" (blank, no value)   │
│                                       │ "제조소(영업소) 명칭" (11pt) │ "" (blank)              │
│                                       │ "변경신청 제품명" (11pt)     │ "" (blank)              │
─────────────────────────────────────────────────────

- Left item names (3 rows): "성명", "제조소(영업소) 명칭", "변경신청 제품명"
- Data: All display as blank, fixed.
- Table structure: 3 rows, 2 columns, cell merge/width/font size match the template.
- **Code automation:** No data input, direct fixed output (ex: `st.table` / direct string in Word table).

─────────────────────────────────────────────────────────
│ 2. 변경유형 (12pt)   │ 3. 신청 유형(AR, IR, Cmin, Cmaj 중 선택) (12pt) │
─────────────────────────────────────────────────────────
│ {step7_results[title_key][idx]['title_text']} (11pt)        │ {step7_results[title_key][idx]['output_1_tag']} (11pt)        │
─────────────────────────────────────────────────────────

- **Left (2. 변경유형):**
    - Automated value:  
      `step7_results[title_key][idx]['title_text']`
        (ex: "3.2.P.1 완제의약품의 성상 및 조성\n9. 완제의약품(고형제제 제외)의 조성 변경")
- **Right (3. 신청유형):**
    - First line:  
      `step7_results[title_key][idx]['output_1_tag']`
        (ex: "Cmaj", "AR", etc.)
- **Table structure:** 2 rows, 2 columns (actually one row, line break in each cell), template width/spacing/font.

────────────────────────────────────────────────────────────
│ 4. 충족조건 (12pt)          │ 조건 충족 여부(○, X 중 선택) (12pt)      │
────────────────────────────────────────────────────────────
│ {step6_items[title_key]['requirements'][k]} (11pt) │ {if step6_selections[f"{title_key}_req_{k}"] == '충족': display = '○'; if step6_selections[f"{title_key}_req_{k}"] == '미충족': display = '×'} (11pt) │
│ ... (repeat for each requirement, auto-expand) ...        │  ... (repeat for each requirement, auto-expand, same logic) ...│
────────────────────────────────────────────────────────────

- **Left:** requirements item (ex: "1. ~조건명"), auto-generated row for each item
    - Actual code:  
      ```python
      for k, v in step6_items[title_key]['requirements'].items():
          left = v
          if step6_selections[f"{title_key}_req_{k}"] == '충족':
              right = '○'
          if step6_selections[f"{title_key}_req_{k}"] == '미충족':
              right = '×'
      ```
- **Table structure:** n rows, 2 columns, 11pt, header row 12pt

────────────────────────────────────────────────────────────
│ 5. 필요서류 (해당 필요서류 기재) (12pt)      │ 구비 여부 (○, X 중 선택) / 해당 페이지 표시 (12pt) │
────────────────────────────────────────────────────────────
│ {each line: split output_2_text by line break} (11pt) │ (always blank, auto)                                           │
│ ... repeat for number of required documents ...        │ ...                                                         │
────────────────────────────────────────────────────────────
    - "필요서류는 다음과 같습니다.'\n" 부분은 첫번째 행에 도출되지 않게 해주세요.

- **Left:**  
    - For the current page combination  
      `step7_results[title_key][idx]['output_2_text']`  
    - Split by '\n', each line = 1 row (11pt)
- **Right:**  
    - Always blank, no direct input
- **Table structure:** n rows, 2 columns, as many rows as required documents, header row 12pt

---

### Code Key/Path-Based Specification (Template 100% Match)

─────────────────────────────────────────────
│ [파일 다운로드]                     [인쇄하기] │ ← Streamlit/Docx identical, in Korean only
│ 「의약품 허가 후 제조방법 변경관리 가이드라인(민원인 안내서)」[붙임] 신청양식 예시 │
│                   {current_page+1} / {total_pages}                    │
─────────────────────────────────────────────

─────────────────────────────────────────────────────
│ 1. 신청인 (vertical merge of 3 rows, 12pt, Bold) │ "성명" (11pt)        │ "" (blank, no value)   │
│                                       │ "제조소(영업소) 명칭" (11pt) │ "" (blank)              │
│                                       │ "변경신청 제품명" (11pt)     │ "" (blank)              │
─────────────────────────────────────────────────────
# Example code:
applicant_table = [
    ["1. 신청인", "성명", ""],
    ["", "제조소(영업소) 명칭", ""],
    ["", "변경신청 제품명", ""]
]
# Column width/row height/font must match docx, including in Streamlit

─────────────────────────────────────────────────────────────
│ 2. 변경유형 (12pt, Bold)                                   │ 3. 신청유형(AR, IR, Cmin, Cmaj 중 선택) (12pt, Bold)        │
│───────────────────────────────────────────────────────────│
│ step7_results[title_key][idx]['title_text']   (11pt)      │ step7_results[title_key][idx]['output_1_tag'] (11pt)       │
─────────────────────────────────────────────────────────────

# Code:
# Left: step7_results[title_key][idx]['title_text']
# Right 1st row: step7_results[title_key][idx]['output_1_tag']

─────────────────────────────────────────────────────────────
│ 4. 충족조건 (12pt, Bold)                                         │ 조건 충족 여부 (○, X, 12pt, Bold)        │
│─────────────────────────────────────────────────────────────────│
│ for k, v in step6_items[title_key]['requirements'].items(): │
│   v (11pt)                                                      │ if step6_selections[f"{title_key}_req_{k}"] == '충족': display = '○'; if step6_selections[f"{title_key}_req_{k}"] == '미충족': display = '×' (11pt) │
│ ... (auto-expand for number of requirements)                     │ ... (repeat for all, auto-expand) ...│
─────────────────────────────────────────────────────────────

# Code:
# for k, v in step6_items[title_key]['requirements'].items():
#    left = v
#    if step6_selections[f"{title_key}_req_{k}"] == '충족':
#        right = '○'
#    if step6_selections[f"{title_key}_req_{k}"] == '미충족':
#        right = '×'

─────────────────────────────────────────────────────────────
│ 5. 필요서류 (12pt, Bold)      │ 구비 여부 (○, X, 12pt, Bold) / 해당 페이지 표시    │
│───────────────────────────────────────────────────────────│
│ for line in step7_results[title_key][idx]['output_2_text'].split('\n'): │
│   line (11pt)                                            │   "" (blank, auto)                    │
│ ... (auto-expand for number of required documents)        │
─────────────────────────────────────────────────────────────
    - "필요서류는 다음과 같습니다.'\n" 부분은 첫번째 행에 도출되지 않게 해주세요.

# Code:
# for line in step7_results[title_key][idx]['output_2_text'].split('\n'):
#     left = line
#     right = ""

─────────────────────────────────────────────────────────────
# For every page and combination (75 cases), mapping is as follows:
# - title_key, idx: determined by page_list[page index]
# - step7_results[title_key][idx]: result for each combination
# - step6_items[title_key]['requirements']: requirements list
# - step6_selections[f"{title_key}_req_{k}"]: '충족'/'미충족'
# Table style, merge, column width, font, alignment, spacing: match exactly to the template docx.

─────────────────────────────────────────────────────────────
# Example function (auto table rendering per page – Streamlit/Python)

def render_applicant_table():
    return [
        ["1. 신청인", "성명", ""],
        ["", "제조소(영업소) 명칭", ""],
        ["", "변경신청 제품명", ""]
    ]

def render_change_type_table(step7_results, title_key, idx):
    return [
        ["2. 변경유형", "3. 신청유형(AR, IR, Cmin, Cmaj 중 선택)"],
        [
            step7_results[title_key][idx]['title_text'],
            f"{step7_results[title_key][idx]['output_1_tag']}"
        ]
    ]

def render_requirements_table(step6_items, step6_selections, title_key):
    rows = []
    for k, v in step6_items[title_key]['requirements'].items():
        if step6_selections[f"{title_key}_req_{k}"] == '충족':
            mark = '○'
        elif step6_selections[f"{title_key}_req_{k}"] == '미충족':
            mark = '×'
        else:
            mark = ""
        rows.append([v, mark])
    return [["4. 충족조건", "조건 충족 여부"]] + rows

def render_documents_table(step7_results, title_key, idx):
    lines = step7_results[title_key][idx]['output_2_text'].split('\n')
    rows = [[line, ""] for line in lines if line.strip()]
    return [["5. 필요서류 (해당하는 필요서류 기재)", "구비 여부 (○, X 중 선택) / 해당 페이지 표시"]] + rows

─────────────────────────────────────────────────────────────

# All tables/cells/values must be accessed via the above keys and variables in the code (no freeform description allowed).
# Word file (.docx) and Streamlit screen must match structure/content/style 100%.

─────────────────────────────────────────────────────────────

# Summary and Notes
- No freeform explanation or abstract description.
- All table/cell/content/automation keys, paths, structures, and code examples only.
- Must not miss any microstructure from the template (font, merge, spacing, etc).
- Table/data automation must be implemented separately for each (title_key, idx).

---

## 3. Overall Code/Data Access Summary (for implementation)

- title_key, idx:  
    - page_list = [(title_key, idx), ...]  
    - Use title_key, idx for each page
- step7_results:  
    - {title_key: [ { 'title_text', 'output_1_tag', 'output_2_text' }, ... ] }
- step6_items:  
    - {title_key: { 'requirements': {k: v, ... } } }
- step6_selections:  
    - {f"{title_key}_req_{k}": '충족'/'미충족', ... }

---

## 4. Example (Actual Automated Data Binding for One Page)

Example: title_key="p1_9", idx=0

- 2. 변경유형: step7_results['p1_9'][0]['title_text']
- 3. 신청유형(AR/IR/Cmin/Cmaj): step7_results['p1_9'][0]['output_1_tag']
- 4. 충족조건:
    - for k, v in step6_items['p1_9']['requirements'].items():
        - left: v
        - right: '○' if step6_selections['p1_9_req_' + k] == '충족'  
                  '×' if step6_selections['p1_9_req_' + k] == '미충족'
- 5. 필요서류:
    - for line in step7_results['p1_9'][0]['output_2_text'].split('\n'):
        - left: line
        - right: blank
    - "필요서류는 다음과 같습니다.'\n" 부분은 첫번째 행에 도출되지 않게 해주세요.

---

## 5. Development/Test/Documentation Guideline

- Table/header/row/column width/font/merge/spacing: 100% match to [제조방법변경 신청양식_empty_.docx] in both UI and Word
- Automation mapping/code/key names: refer to above table, code, variable, and structure, and annotate clearly
- Combination mapping: repeat above structure per (title_key, idx)
- Any omission, error, or mismatch means the work is NOT accepted

---

# All items above MUST be implemented in code, table, and data mapping, with NO omission in development, test, code review, or documentation.

Summary:  
All mapping, code path, cell structure, and data input mode above MUST  
produce a **single table** in which every item is connected vertically in one structure,  
as below:

─────────────────────────────────────────────────────────────
│ 1. 신청인 (vertical merge of 3 rows, 12pt, Bold) │ 성명 (11pt) │ "" (blank) │
│ │ 제조소(영업소) 명칭 │ "" (blank) │
│ │ 변경신청 제품명 │ "" (blank) │
├───────────────────────────────────────┼───────────────────────┼──────────────────┤
│ 2. 변경유형 (vertical merge of 2 rows, 12pt, Bold) │ 3. 신청 유형(AR, IR, Cmin, Cmaj 중 선택) (12pt, Bold) │
│ step7_results[title_key][idx]['title_text'] (11pt) │ step7_results[title_key][idx]['output_1_tag'] (11pt)│
├───────────────────────────────────────┴───────────────────────┴──────────────────┤
│ 4. 충족조건 (header cell, 12pt, Bold, row merge) │ "충족조건" (11pt, Bold) │ "조건 충족 여부(○, X 중 선택)" (11pt, Bold) │
│ for k, v in step6_items[title_key]['requirements'].items(): │ v (11pt) │ '○' if step6_selections[f"{title_key}_req_{k}"] == '충족'  
'×' if step6_selections[f"{title_key}_req_{k}"] == '미충족' (11pt) │
├─────────────────────────────────────────────────────────────┼──────────────────┤
│ 5. 필요서류 (header cell, 12pt, Bold, row merge) │ "필요서류 (해당하는 필요서류 기재)" (11pt, Bold) │ "구비 여부 (○, X 중 선택)" (11pt, Bold) │
│ for line in step7_results[title_key][idx]['output_2_text'].split('\n'): │ line (11pt) │ "" (blank) │
─────────────────────────────────────────────────────────────
    - "필요서류는 다음과 같습니다.'\n" 부분은 첫번째 행에 도출되지 않게 해주세요.

In this way, within **one table**,  
all cell merging, font, alignment, column width/row height,  
data source (code path/key),  
row expansion for 충족조건 and 필요서류 (dynamic),
    - "필요서류는 다음과 같습니다.'\n" 부분은 첫번째 행에 도출되지 않게 해주세요.

must be exactly identical to [제조방법변경 신청양식_empty_.docx],  
and **splitting into multiple tables, or rendering the title separate from the table, is absolutely prohibited.**
