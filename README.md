# NMD RSS Web Part

## Summary

A SharePoint Framework web part that displays RSS feeds in an attractive, responsive card layout. This web part fetches RSS feeds from any public RSS URL and presents the content in a modern, user-friendly interface with support for images, categories, metadata, and pagination.

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.20.0-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

- SharePoint Online or SharePoint 2019/2016 on-premises
- Valid RSS feed URL that is publicly accessible
- Modern SharePoint pages (not classic pages)

## Solution

| Solution | Author(s)            |
| -------- | -------------------- |
| nmd-rss  | NMD Development Team |

## Version history

| Version | Date           | Comments                                                         |
| ------- | -------------- | ---------------------------------------------------------------- |
| 1.1.0   | July 4, 2025   | Current version - Production with responsive 6 or 8 cards layout |
| 1.0.0   | March 18, 2025 | Initial release - Prototype                                      |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**

> Include any additional steps as needed.

## Features

The NMD RSS Web Part provides a comprehensive RSS feed display solution with the following capabilities:

### Core Functionality

- **RSS Feed Parsing**: Fetches and parses RSS feeds from any public URL
- **Responsive Design**: Automatically adapts to different screen sizes and container widths
- **Card-based Layout**: Displays feed items in attractive document cards with consistent styling
- **Pagination**: Supports multiple pages when feed contains many items

### Content Display

- **Article Titles**: Clickable titles that link to the original articles
- **Article Descriptions**: Truncated descriptions (120 characters) with ellipsis
- **Article Images**: Displays featured images with loading spinners and fallbacks
- **Image Credits**: Shows photo credits when available in the RSS feed
- **Categories/Tags**: Displays article categories as styled tags
- **Metadata**: Shows author names and publication dates in localized format

### Channel Information

- **Channel Logo**: Displays the RSS feed's channel image/logo when available
- **Customizable Title**: Configurable web part title independent of feed title

### User Experience

- **Loading States**: Shows spinner while fetching feed data
- **Error Handling**: Graceful error messages for failed feed requests
- **Image Preloading**: Optimized image loading with batching for better performance
- **Responsive Pagination**: Navigation controls that appear only when needed

### Configuration Options

- **RSS Feed URL**: Configurable feed URL (supports both http and https)
- **Web Part Title**: Customizable display title for the web part

This extension illustrates the following concepts:

- RSS feed parsing and XML processing
- Responsive design with dynamic layout adjustments
- Image handling and optimization
- SharePoint Framework property pane configuration
- Modern React patterns with hooks and functional components
- Error handling and loading states
- Accessibility considerations

> Notice that better pictures and documentation will increase the sample usage and the value you are providing for others. Thanks for your submissions advance.

> Share your web part with others through Microsoft 365 Patterns and Practices program to get visibility and exposure. More details on the community, open-source projects and other activities from http://aka.ms/m365pnp.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
