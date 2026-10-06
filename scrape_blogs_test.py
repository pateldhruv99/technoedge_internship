import time
import requests
import pandas as pd
from bs4 import BeautifulSoup


BASE_URL = "https://www.technoedgels.com/blog-insights/"
OUTPUT_FILE = "test_blogs.xlsx"

# Test only first 3 articles
ARTICLE_LIMIT = 3

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/154.0 Safari/537.36"
    )
}


def clean_text(text):
    if not text:
        return ""

    return " ".join(text.split())


def get_soup(url):
    print(f"Opening: {url}")

    response = requests.get(
        url,
        headers=HEADERS,
        timeout=30
    )

    response.raise_for_status()

    return BeautifulSoup(
        response.text,
        "lxml"
    )


def extract_cards():
    print("\nReading blog archive...\n")

    soup = get_soup(BASE_URL)

    cards = soup.select(".rt-grid-item")

    print(f"Cards found: {len(cards)}")

    results = []

    for card in cards:

        title_element = card.select_one(
            "h3.entry-title a.tpg-post-link"
        )

        if not title_element:
            continue

        date_element = card.select_one(
            ".post-meta-tags .date a"
        )

        excerpt_element = card.select_one(
            ".tpg-excerpt-inner"
        )

        image_element = card.select_one(
            ".rt-img-holder img.rt-img-responsive"
        )

        blog_url = title_element.get(
            "href",
            ""
        )

        image_url = ""

        if image_element:
            image_url = image_element.get(
                "src",
                ""
            )

        results.append({
            "Card_Title": clean_text(
                title_element.get_text()
            ),
            "Card_Date": clean_text(
                date_element.get_text()
                if date_element
                else ""
            ),
            "Excerpt": clean_text(
                excerpt_element.get_text()
                if excerpt_element
                else ""
            ),
            "Thumbnail_URL": image_url,
            "Blog_URL": blog_url
        })

    return results


def extract_article(blog_url):

    soup = get_soup(blog_url)

    article = soup.select_one(
        "article.ast-article-single"
    )

    if not article:
        raise Exception(
            "Article container not found"
        )

    title_element = article.select_one(
        ".entry-header h1.entry-title"
    )

    author_element = article.select_one(
        ".entry-meta .author-name"
    )

    date_element = article.select_one(
        ".entry-meta .published"
    )

    body = article.select_one(
        ".entry-content.clear"
    )

    if not body:
        raise Exception(
            "Article body not found"
        )

    # Extract H2 headings
    h2_headings = []

    for heading in body.select(
        "h2.wp-block-heading"
    ):
        h2_headings.append(
            clean_text(
                heading.get_text()
            )
        )

    # Extract H3 headings
    h3_headings = []

    for heading in body.select(
        "h3.wp-block-heading"
    ):
        h3_headings.append(
            clean_text(
                heading.get_text()
            )
        )

    # Extract clean article text
    content_parts = []

    for element in body.find_all(
        ["h2", "h3", "p", "li"]
    ):

        text = clean_text(
            element.get_text()
        )

        if not text:
            continue

        if element.name == "h2":
            content_parts.append(
                f"\n## {text}\n"
            )

        elif element.name == "h3":
            content_parts.append(
                f"\n### {text}\n"
            )

        elif element.name == "li":
            content_parts.append(
                f"- {text}"
            )

        else:
            content_parts.append(
                text
            )

    full_article = "\n\n".join(
        content_parts
    )

    return {
        "Article_Title": clean_text(
            title_element.get_text()
            if title_element
            else ""
        ),

        "Author": clean_text(
            author_element.get_text()
            if author_element
            else ""
        ),

        "Published_Date": clean_text(
            date_element.get_text()
            if date_element
            else ""
        ),

        "H2_Headings": " | ".join(
            h2_headings
        ),

        "H3_Headings": " | ".join(
            h3_headings
        ),

        "Full_Article": full_article
    }


def main():

    print(
        "\n=============================="
    )

    print(
        "TechnoEdge Blog Scraper TEST"
    )

    print(
        "==============================\n"
    )

    cards = extract_cards()

    if not cards:
        print(
            "No blog cards were found."
        )

        return

    print(
        f"\nTotal cards detected: "
        f"{len(cards)}"
    )

    print(
        f"Testing first "
        f"{ARTICLE_LIMIT} articles.\n"
    )

    final_results = []

    for index, card in enumerate(
        cards[:ARTICLE_LIMIT],
        start=1
    ):

        print(
            f"\n------------------------"
        )

        print(
            f"Article {index}/"
            f"{ARTICLE_LIMIT}"
        )

        print(
            card["Card_Title"]
        )

        try:

            article_data = extract_article(
                card["Blog_URL"]
            )

            card.update(
                article_data
            )

            card["Status"] = "OK"

            card["Error"] = ""

            print(
                "SUCCESS"
            )

        except Exception as error:

            card["Article_Title"] = ""
            card["Author"] = ""
            card["Published_Date"] = ""
            card["H2_Headings"] = ""
            card["H3_Headings"] = ""
            card["Full_Article"] = ""

            card["Status"] = "FAILED"

            card["Error"] = str(error)

            print(
                f"FAILED: {error}"
            )

        final_results.append(
            card
        )

        # Small delay between requests
        time.sleep(1)

    dataframe = pd.DataFrame(
        final_results
    )

    dataframe.insert(
        0,
        "Sr_No",
        range(
            1,
            len(dataframe) + 1
        )
    )

    dataframe.to_excel(
        OUTPUT_FILE,
        index=False
    )

    print(
        "\n=============================="
    )

    print(
        "TEST COMPLETED"
    )

    print(
        "=============================="
    )

    print(
        f"\nExcel created: "
        f"{OUTPUT_FILE}"
    )


if __name__ == "__main__":
    main()