import tweepy
from openpyxl import Workbook
from openpyxl import load_workbook

api_key = ""
api_secret_key = ""
bearer_token = ""
access_token = ""
access_token_secret = ""

auth = tweepy.OAuthHandler(api_key, api_secret_key)
auth.set_access_token(access_token, access_token_secret)
api = tweepy.API(auth)

# tweets = api.search(q="sustainability", lang="en", count=10)
# tweets = api.search(q="sustainability -filter:retweets", lang="en", until="2021-05-21", tweet_mode='extended', count=10)

# q = "climate change"
# q = "carbon footprint"
q = "csr"

tweets = api.search_30_day(environment_name="sasascraper30days", query=q + " h&m lang:en")

tweets_list = []

for tweet in tweets:
    # tweets_list.append(tweet.full_text)
    # print(tweet.full_text)
    if tweet.truncated:
        tweets_list.append(tweet.extended_tweet['full_text'])
        print(tweet.extended_tweet['full_text'])

wb = load_workbook('Dataset2.xlsx')
ws = wb.active

for tweet in tweets_list:
    ws.append([tweet])

wb.save('Dataset2.xlsx')
