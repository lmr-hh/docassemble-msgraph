# Docassemble-MSGraph

This project is a wrapper around the popular `requests_oauthlib` providing
a simple interface to interact with Microsoft Graph APIs in docassemble.

## Usage
The library provides a simple interface that can be instaciated as follows:

```python
from docassemble.msgraph import MSGraphSession

graph = MSGraphSession()
graph.fetch_token()
```

By default this will read the `tenant id`, `client id`, and `client secret` from
a `msgraph` section in the docassemble configuration. This can be customized
however. See the `MSGraphSession` class for details.

The `MSGraphSession` class also provides several useful methods for typical
tasks.

## Debugging

You can supply values for `tenant_id`, `client_id` and `client_secret` directly
instead of using the docassemble configuration. This may be useful when testing
your code locally.