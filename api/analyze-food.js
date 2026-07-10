const VALID_TAGS = ["탄수화물 많음", "균형 좋음", "비타민 부족", "가공식품 많음", "단백질 보완"];
const VALID_LEVELS = ["낮음", "보통", "높음"];

const PROMPT = `당신은 친절한 식단 코치입니다. 첨부된 음식 사진 하나를 보고 아래 JSON 형식으로만 답하세요.
다른 설명, 마크다운, 코드블록 없이 순수 JSON 객체 하나만 출력하세요.

{
  "food": "사진 속 음식 이름 (예: 뼈해장국, 샐러드, 김치찌개 등). 정확한 이름을 모르면 보이는 재료로 최대한 구체적으로 추정",
  "탄수화물": "낮음" | "보통" | "높음",
  "단백질": "낮음" | "보통" | "높음",
  "지방": "낮음" | "보통" | "높음",
  "채소/비타민": "낮음" | "보통" | "높음",
  "tag": "탄수화물 많음" | "균형 좋음" | "비타민 부족" | "가공식품 많음" | "단백질 보완",
  "comment": "food 이름을 자연스럽게 언급하며 한국어 한 문장, 친근하고 다정한 말투의 피드백"
}`;

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Method not allowed" });
    return;
  }

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    res.status(500).json({ error: "GEMINI_API_KEY not configured" });
    return;
  }

  const { image } = req.body || {};
  const match = typeof image === "string" && image.match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);
  if (!match) {
    res.status(400).json({ error: "Invalid image data" });
    return;
  }
  const [, mimeType, base64Data] = match;

  try {
    const geminiRes = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`,
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          contents: [
            {
              parts: [{ text: PROMPT }, { inline_data: { mime_type: mimeType, data: base64Data } }],
            },
          ],
        }),
      }
    );

    const json = await geminiRes.json();
    if (!geminiRes.ok) {
      console.error("Gemini API 오류 응답:", JSON.stringify(json));
      res.status(502).json({ error: json?.error?.message || "Gemini API error" });
      return;
    }

    const text = json?.candidates?.[0]?.content?.parts?.[0]?.text || "";
    const braceMatch = text.match(/\{[\s\S]*\}/);
    if (!braceMatch) {
      console.error("Gemini 응답에서 JSON을 찾지 못함. 원본 텍스트:", text);
      throw new Error("No JSON in response");
    }
    const parsed = JSON.parse(braceMatch[0]);

    if (
      !VALID_TAGS.includes(parsed.tag) ||
      !VALID_LEVELS.includes(parsed["탄수화물"]) ||
      typeof parsed.food !== "string" ||
      !parsed.food.trim()
    ) {
      console.error("Gemini 응답 형식이 예상과 다름:", JSON.stringify(parsed));
      throw new Error("Unexpected response shape");
    }

    res.status(200).json(parsed);
  } catch (err) {
    console.error("음식 사진 분석 실패:", err?.message || err);
    res.status(502).json({ error: "Analysis failed" });
  }
}
