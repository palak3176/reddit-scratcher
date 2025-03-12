import praw
from textblob import TextBlob
import pandas as pd

# Initialize Reddit API
reddit = praw.Reddit(
client_id='',
client_secret='',
user_agent='',
username='',
password='',
)

def get_reddit_posts(company_name):
    posts = reddit.subreddit('all').search(company_name, limit=100)
    return posts

def analyze_sentiment(post):
    analysis = TextBlob(post.title)
    if analysis.sentiment.polarity > 0:
        return 'Positive'
    elif analysis.sentiment.polarity < 0:
        return 'Negative'
    else:
        return 'Neutral'

def summarize_posts(company_name):
    posts = get_reddit_posts(company_name)
    summary = {'Positive': 0, 'Negative': 0, 'Neutral': 0}
    detailed_summary = []

    for post in posts:
        sentiment = analyze_sentiment(post)
        detailed_summary.append({
            'Title': post.title,
            'Sentiment': sentiment,
            'Score': post.score,
            'URL': post.url
        })
        summary[sentiment] += 1
    
    return summary, detailed_summary

def export_to_excel(detailed_posts, company_name):
    df = pd.DataFrame(detailed_posts)
    df.to_excel(f"{company_name}_reddit_sentiment_analysis.xlsx", index=False)

if __name__ == "__main__":
    company_name = input("Enter the company name: ")
    sentiment_summary, detailed_posts = summarize_posts(company_name)
    
    print(f"Sentiment Summary for {company_name}:")
    print(f"Positive: {sentiment_summary['Positive']}")
    print(f"Negative: {sentiment_summary['Negative']}")
    print(f"Neutral: {sentiment_summary['Neutral']}")
    
    print("\nDetailed Posts Summary:")
    for post in detailed_posts:
        print(f"Title: {post['Title']}, Sentiment: {post['Sentiment']}, Score: {post['Score']}, URL: {post['URL']}")
    
    export_to_excel(detailed_posts, company_name)
    print(f"\nData exported to {company_name}_reddit_sentiment_analysis.xlsx")