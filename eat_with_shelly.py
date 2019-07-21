""" Quickstart script for InstaPy usage """

# imports
from instapy import InstaPy
from instapy import smart_run
from instapy import set_workspace


# set workspace folder at desired location (default is at your home folder)
set_workspace(path=None)

# get an InstaPy session!
insta_username = 'eat_with_shelly'
insta_password = 'Miao0902'
session = InstaPy(username=insta_username,
                  password=insta_password,
                  headless_browser=False)

with smart_run(session):
    # settings
    session.set_relationship_bounds(enabled=True,
				 potency_ratio=None,
				  delimit_by_numbers=True,
				   max_followers=4590,
				    max_following=5555,
				     min_followers=5,
				      min_following=7)
    session.set_do_comment(True, percentage=80)
    session.set_comments(['Amazing!', 'So good!!', 'Nice!', 'Nice pic!', 'Good picture!', 'Excellent!', 'Looks good!', 'Wonderful!'])
    session.set_dont_include(['kuilinchen', 'aoimiao'])
    session.set_dont_like(['nsfw'])

    # actions
    #session.set_do_follow(enabled=True, percentage=20, times=2)
    #session.unfollow_users(amount=126, nonFollowers=True, style="RANDOM", unfollow_after=42*60*60, sleep_delay=501)
    session.like_by_tags(['food'], amount=500)
