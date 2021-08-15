import tweepy
import pandas as pd

api_key = ""
api_secret_key = ""
bearer_token = ""
access_token = ""
access_token_secret = ""

auth = tweepy.OAuthHandler(api_key, api_secret_key)
auth.set_access_token(access_token, access_token_secret)
api = tweepy.API(auth)

tweets = api.user_timeline(id="hm", count=200, tweet_mode='extended')

idlist = []
tweetlist = []
favlist = []
retweetlist = []

for tweet in tweets:
    if (not tweet.retweeted) and ('RT @' not in tweet.full_text):
        idlist.append(tweet.id_str)
        tweetlist.append(tweet.full_text)
        favlist.append(tweet.favorite_count)
        retweetlist.append(tweet.retweet_count)

tweetdict = {'id': idlist, 'tweet': tweetlist, 'fav_count': favlist, 'retweet_count': retweetlist}

df = pd.DataFrame(tweetdict)
df.to_excel('newhmscraper.xlsx', index=False)
