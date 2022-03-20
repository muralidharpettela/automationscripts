import json

from facebook import GraphAPI


def read_creds(filename):
    '''
    Store API credentials in a safe place.
    If you use Git, make sure to add the file to .gitignore
    '''
    with open(filename) as f:
        credentials = json.load(f)
    return credentials


credentials = read_creds('credentials.json')

graph = GraphAPI(access_token=credentials['access_token'])

message = '''
Add your message here.
'''
link = 'https://www.jcchouinard.com/'
groups = ['744128789503859']

for group in groups:
    graph.put_object(group, 'feed', message=message, link=link)
    print(graph.get_connections(group, 'feed'))