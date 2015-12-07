



<!DOCTYPE html>
<html lang="en" class=" is-copy-enabled">
  <head prefix="og: http://ogp.me/ns# fb: http://ogp.me/ns/fb# object: http://ogp.me/ns/object# article: http://ogp.me/ns/article# profile: http://ogp.me/ns/profile#">
    <meta charset='utf-8'>
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta http-equiv="Content-Language" content="en">
    <meta name="viewport" content="width=1020">
    
    
    <title>Apricot/Weather.py at master · kawatan/Apricot</title>
    <link rel="search" type="application/opensearchdescription+xml" href="/opensearch.xml" title="GitHub">
    <link rel="fluid-icon" href="https://github.com/fluidicon.png" title="GitHub">
    <link rel="apple-touch-icon" sizes="57x57" href="/apple-touch-icon-114.png">
    <link rel="apple-touch-icon" sizes="114x114" href="/apple-touch-icon-114.png">
    <link rel="apple-touch-icon" sizes="72x72" href="/apple-touch-icon-144.png">
    <link rel="apple-touch-icon" sizes="144x144" href="/apple-touch-icon-144.png">
    <meta property="fb:app_id" content="1401488693436528">

      <meta content="@github" name="twitter:site" /><meta content="summary" name="twitter:card" /><meta content="kawatan/Apricot" name="twitter:title" /><meta content="Apricot - Desktop Mascot for Windows" name="twitter:description" /><meta content="https://avatars0.githubusercontent.com/u/246691?v=3&amp;s=400" name="twitter:image:src" />
      <meta content="GitHub" property="og:site_name" /><meta content="object" property="og:type" /><meta content="https://avatars0.githubusercontent.com/u/246691?v=3&amp;s=400" property="og:image" /><meta content="kawatan/Apricot" property="og:title" /><meta content="https://github.com/kawatan/Apricot" property="og:url" /><meta content="Apricot - Desktop Mascot for Windows" property="og:description" />
      <meta name="browser-stats-url" content="https://api.github.com/_private/browser/stats">
    <meta name="browser-errors-url" content="https://api.github.com/_private/browser/errors">
    <link rel="assets" href="https://assets-cdn.github.com/">
    <link rel="web-socket" href="wss://live.github.com/_sockets/OTAzOTQxNzo2OGViMmNiMDM4Mjc5YTIzMGZkYjAwYzI0ZTIyZTcxNToyY2RjMTM1MWMzYmI0MjlkYmQxZjg5M2MyYjVmZTZhN2FhN2YxMTg3OWI0YTRhMjQ4MmRlN2E0NDVlYTEzYTNl--3fe1b7b18bcd7bf76d0bfd777f7fd43b2ee2732a">
    <meta name="pjax-timeout" content="1000">
    <link rel="sudo-modal" href="/sessions/sudo_modal">

    <meta name="msapplication-TileImage" content="/windows-tile.png">
    <meta name="msapplication-TileColor" content="#ffffff">
    <meta name="selected-link" value="repo_source" data-pjax-transient>

    <meta name="google-site-verification" content="KT5gs8h0wvaagLKAVWq8bbeNwnZZK1r1XQysX3xurLU">
    <meta name="google-analytics" content="UA-3769691-2">

<meta content="collector.githubapp.com" name="octolytics-host" /><meta content="github" name="octolytics-app-id" /><meta content="7E75FABD:4082:6AE3402:56659B93" name="octolytics-dimension-request_id" /><meta content="9039417" name="octolytics-actor-id" /><meta content="YA-androidapp" name="octolytics-actor-login" /><meta content="7fe57715b9d8b23fe1abeb30a97b0a2cc2e3d59fb079e9f3df2133f42b7469e4" name="octolytics-actor-hash" />
<meta content="/&lt;user-name&gt;/&lt;repo-name&gt;/blob/show" data-pjax-transient="true" name="analytics-location" />
<meta content="Rails, view, blob#show" data-pjax-transient="true" name="analytics-event" />


  <meta class="js-ga-set" name="dimension1" content="Logged In">



        <meta name="hostname" content="github.com">
    <meta name="user-login" content="YA-androidapp">

        <meta name="expected-hostname" content="github.com">

      <link rel="mask-icon" href="https://assets-cdn.github.com/pinned-octocat.svg" color="#4078c0">
      <link rel="icon" type="image/x-icon" href="https://assets-cdn.github.com/favicon.ico">

    <meta content="283793ba78ae1eabcbd2e2c40a1b786ff2ce6e7e" name="form-nonce" />

    <link crossorigin="anonymous" href="https://assets-cdn.github.com/assets/github-980086798633a7804d9e6353689ba4c79120d0f4a6264214d08411a5f3c6116a.css" media="all" rel="stylesheet" />
    <link crossorigin="anonymous" href="https://assets-cdn.github.com/assets/github2-513b3125a367b7521d4c9c18f9f9fb26186febfcccecc45164cbffd3c544db1b.css" media="all" rel="stylesheet" />
    
    
    


    <meta http-equiv="x-pjax-version" content="fb6d8f209402822d6fc2b9a3f89e2d36">

      
  <meta name="description" content="Apricot - Desktop Mascot for Windows">
  <meta name="go-import" content="github.com/kawatan/Apricot git https://github.com/kawatan/Apricot.git">

  <meta content="246691" name="octolytics-dimension-user_id" /><meta content="kawatan" name="octolytics-dimension-user_login" /><meta content="693579" name="octolytics-dimension-repository_id" /><meta content="kawatan/Apricot" name="octolytics-dimension-repository_nwo" /><meta content="true" name="octolytics-dimension-repository_public" /><meta content="false" name="octolytics-dimension-repository_is_fork" /><meta content="693579" name="octolytics-dimension-repository_network_root_id" /><meta content="kawatan/Apricot" name="octolytics-dimension-repository_network_root_nwo" />
  <link href="https://github.com/kawatan/Apricot/commits/master.atom" rel="alternate" title="Recent Commits to Apricot:master" type="application/atom+xml">

  </head>


  <body class="logged_in   env-production windows vis-public page-blob">
    <a href="#start-of-content" tabindex="1" class="accessibility-aid js-skip-to-content">Skip to content</a>

    
    
    



      <div class="header header-logged-in true" role="banner">
  <div class="container clearfix">

    <a class="header-logo-invertocat" href="https://github.com/" data-hotkey="g d" aria-label="Homepage" data-ga-click="Header, go to dashboard, icon:logo">
  <span class="mega-octicon octicon-mark-github"></span>
</a>


      <div class="site-search repo-scope js-site-search" role="search">
          <!-- </textarea> --><!-- '"` --><form accept-charset="UTF-8" action="/kawatan/Apricot/search" class="js-site-search-form" data-global-search-url="/search" data-repo-search-url="/kawatan/Apricot/search" method="get"><div style="margin:0;padding:0;display:inline"><input name="utf8" type="hidden" value="&#x2713;" /></div>
  <label class="js-chromeless-input-container form-control">
    <div class="scope-badge">This repository</div>
    <input type="text"
      class="js-site-search-focus js-site-search-field is-clearable chromeless-input"
      data-hotkey="s"
      name="q"
      placeholder="Search"
      aria-label="Search this repository"
      data-global-scope-placeholder="Search GitHub"
      data-repo-scope-placeholder="Search"
      tabindex="1"
      autocapitalize="off">
  </label>
</form>
      </div>

      <ul class="header-nav left" role="navigation">
        <li class="header-nav-item">
          <a href="/pulls" class="js-selected-navigation-item header-nav-link" data-ga-click="Header, click, Nav menu - item:pulls context:user" data-hotkey="g p" data-selected-links="/pulls /pulls/assigned /pulls/mentioned /pulls">
            Pull requests
</a>        </li>
        <li class="header-nav-item">
          <a href="/issues" class="js-selected-navigation-item header-nav-link" data-ga-click="Header, click, Nav menu - item:issues context:user" data-hotkey="g i" data-selected-links="/issues /issues/assigned /issues/mentioned /issues">
            Issues
</a>        </li>
          <li class="header-nav-item">
            <a class="header-nav-link" href="https://gist.github.com/" data-ga-click="Header, go to gist, text:gist">Gist</a>
          </li>
      </ul>

    
<ul class="header-nav user-nav right" id="user-links">
  <li class="header-nav-item">
      <span class="js-socket-channel js-updatable-content"
        data-channel="notification-changed:YA-androidapp"
        data-url="/notifications/header">
      <a href="/notifications" aria-label="You have unread notifications" class="header-nav-link notification-indicator tooltipped tooltipped-s" data-ga-click="Header, go to notifications, icon:unread" data-hotkey="g n">
          <span class="mail-status unread"></span>
          <span class="octicon octicon-bell"></span>
</a>  </span>

  </li>

  <li class="header-nav-item dropdown js-menu-container">
    <a class="header-nav-link tooltipped tooltipped-s js-menu-target" href="/new"
       aria-label="Create new…"
       data-ga-click="Header, create new, icon:add">
      <span class="octicon octicon-plus left"></span>
      <span class="dropdown-caret"></span>
    </a>

    <div class="dropdown-menu-content js-menu-content">
      <ul class="dropdown-menu dropdown-menu-sw">
        
<a class="dropdown-item" href="/new" data-ga-click="Header, create new repository">
  New repository
</a>


  <a class="dropdown-item" href="/organizations/new" data-ga-click="Header, create new organization">
    New organization
  </a>



  <div class="dropdown-divider"></div>
  <div class="dropdown-header">
    <span title="kawatan/Apricot">This repository</span>
  </div>
    <a class="dropdown-item" href="/kawatan/Apricot/issues/new" data-ga-click="Header, create new issue">
      New issue
    </a>

      </ul>
    </div>
  </li>

  <li class="header-nav-item dropdown js-menu-container">
    <a class="header-nav-link name tooltipped tooltipped-sw js-menu-target" href="/YA-androidapp"
       aria-label="View profile and more"
       data-ga-click="Header, show menu, icon:avatar">
      <img alt="@YA-androidapp" class="avatar" height="20" src="https://avatars0.githubusercontent.com/u/9039417?v=3&amp;s=40" width="20" />
      <span class="dropdown-caret"></span>
    </a>

    <div class="dropdown-menu-content js-menu-content">
      <div class="dropdown-menu  dropdown-menu-sw">
        <div class=" dropdown-header header-nav-current-user css-truncate">
            Signed in as <strong class="css-truncate-target">YA-androidapp</strong>

        </div>


        <div class="dropdown-divider"></div>

          <a class="dropdown-item" href="/YA-androidapp" data-ga-click="Header, go to profile, text:your profile">
            Your profile
          </a>
        <a class="dropdown-item" href="/stars" data-ga-click="Header, go to starred repos, text:your stars">
          Your stars
        </a>
        <a class="dropdown-item" href="/explore" data-ga-click="Header, go to explore, text:explore">
          Explore
        </a>
          <a class="dropdown-item" href="/integrations" data-ga-click="Header, go to integrations, text:integrations">
            Integrations
          </a>
        <a class="dropdown-item" href="https://help.github.com" data-ga-click="Header, go to help, text:help">
          Help
        </a>

          <div class="dropdown-divider"></div>

          <a class="dropdown-item" href="/settings/profile" data-ga-click="Header, go to settings, icon:settings">
            Settings
          </a>

          <!-- </textarea> --><!-- '"` --><form accept-charset="UTF-8" action="/logout" class="logout-form" data-form-nonce="283793ba78ae1eabcbd2e2c40a1b786ff2ce6e7e" method="post"><div style="margin:0;padding:0;display:inline"><input name="utf8" type="hidden" value="&#x2713;" /><input name="authenticity_token" type="hidden" value="x6SzpXPr5PpzyfvhecLDqRy9yyx3Kw+4L5bJWrujvTO6fzKAxDuUenuKk1IdrcOD5qI3yl0JzW65Nv9Ak8OR5g==" /></div>
            <button class="dropdown-item dropdown-signout" data-ga-click="Header, sign out, icon:logout">
              Sign out
            </button>
</form>
      </div>
    </div>
  </li>
</ul>


    
  </div>
</div>

      

      


    <div id="start-of-content" class="accessibility-aid"></div>

    <div id="js-flash-container">
</div>


    <div role="main" class="main-content">
        <div itemscope itemtype="http://schema.org/WebPage">
    <div id="js-repo-pjax-container" class="context-loader-container js-repo-nav-next" data-pjax-container>
      
<div class="pagehead repohead instapaper_ignore readability-menu experiment-repo-nav">
  <div class="container repohead-details-container">

    

<ul class="pagehead-actions">

  <li>
        <!-- </textarea> --><!-- '"` --><form accept-charset="UTF-8" action="/notifications/subscribe" class="js-social-container" data-autosubmit="true" data-form-nonce="283793ba78ae1eabcbd2e2c40a1b786ff2ce6e7e" data-remote="true" method="post"><div style="margin:0;padding:0;display:inline"><input name="utf8" type="hidden" value="&#x2713;" /><input name="authenticity_token" type="hidden" value="kWwPKYf41VIAS9LTCTjT1pQM/7yd9uccdcaOZt+mbGvXeDkDlcndkc8oA2dZXNuvZ6fRZqWSiVYhDMBW6C7VIg==" /></div>      <input id="repository_id" name="repository_id" type="hidden" value="693579" />

        <div class="select-menu js-menu-container js-select-menu">
          <a href="/kawatan/Apricot/subscription"
            class="btn btn-sm btn-with-count select-menu-button js-menu-target" role="button" tabindex="0" aria-haspopup="true"
            data-ga-click="Repository, click Watch settings, action:blob#show">
            <span class="js-select-button">
              <span class="octicon octicon-eye"></span>
              Unwatch
            </span>
          </a>
          <a class="social-count js-social-count" href="/kawatan/Apricot/watchers">
            4
          </a>

        <div class="select-menu-modal-holder">
          <div class="select-menu-modal subscription-menu-modal js-menu-content" aria-hidden="true">
            <div class="select-menu-header">
              <span class="octicon octicon-x js-menu-close" role="button" aria-label="Close"></span>
              <span class="select-menu-title">Notifications</span>
            </div>

              <div class="select-menu-list js-navigation-container" role="menu">

                <div class="select-menu-item js-navigation-item " role="menuitem" tabindex="0">
                  <span class="select-menu-item-icon octicon octicon-check"></span>
                  <div class="select-menu-item-text">
                    <input id="do_included" name="do" type="radio" value="included" />
                    <span class="select-menu-item-heading">Not watching</span>
                    <span class="description">Be notified when participating or @mentioned.</span>
                    <span class="js-select-button-text hidden-select-button-text">
                      <span class="octicon octicon-eye"></span>
                      Watch
                    </span>
                  </div>
                </div>

                <div class="select-menu-item js-navigation-item selected" role="menuitem" tabindex="0">
                  <span class="select-menu-item-icon octicon octicon octicon-check"></span>
                  <div class="select-menu-item-text">
                    <input checked="checked" id="do_subscribed" name="do" type="radio" value="subscribed" />
                    <span class="select-menu-item-heading">Watching</span>
                    <span class="description">Be notified of all conversations.</span>
                    <span class="js-select-button-text hidden-select-button-text">
                      <span class="octicon octicon-eye"></span>
                      Unwatch
                    </span>
                  </div>
                </div>

                <div class="select-menu-item js-navigation-item " role="menuitem" tabindex="0">
                  <span class="select-menu-item-icon octicon octicon-check"></span>
                  <div class="select-menu-item-text">
                    <input id="do_ignore" name="do" type="radio" value="ignore" />
                    <span class="select-menu-item-heading">Ignoring</span>
                    <span class="description">Never be notified.</span>
                    <span class="js-select-button-text hidden-select-button-text">
                      <span class="octicon octicon-mute"></span>
                      Stop ignoring
                    </span>
                  </div>
                </div>

              </div>

            </div>
          </div>
        </div>
</form>
  </li>

  <li>
    
  <div class="js-toggler-container js-social-container starring-container on">

    <!-- </textarea> --><!-- '"` --><form accept-charset="UTF-8" action="/kawatan/Apricot/unstar" class="js-toggler-form starred js-unstar-button" data-form-nonce="283793ba78ae1eabcbd2e2c40a1b786ff2ce6e7e" data-remote="true" method="post"><div style="margin:0;padding:0;display:inline"><input name="utf8" type="hidden" value="&#x2713;" /><input name="authenticity_token" type="hidden" value="WDeKrQxg9+3rBjSO/h7wmus0HvL82Bc7riGJhgsoyux/2jBha6ayQHri8e9y3AexOW6OgM1YQNKY7Hdc0kMb9Q==" /></div>
      <button
        class="btn btn-sm btn-with-count js-toggler-target"
        aria-label="Unstar this repository" title="Unstar kawatan/Apricot"
        data-ga-click="Repository, click unstar button, action:blob#show; text:Unstar">
        <span class="octicon octicon-star"></span>
        Unstar
      </button>
        <a class="social-count js-social-count" href="/kawatan/Apricot/stargazers">
          10
        </a>
</form>
    <!-- </textarea> --><!-- '"` --><form accept-charset="UTF-8" action="/kawatan/Apricot/star" class="js-toggler-form unstarred js-star-button" data-form-nonce="283793ba78ae1eabcbd2e2c40a1b786ff2ce6e7e" data-remote="true" method="post"><div style="margin:0;padding:0;display:inline"><input name="utf8" type="hidden" value="&#x2713;" /><input name="authenticity_token" type="hidden" value="ZGNoCO6YBg/K4cRF8CX+7NXqmTDPG9HKeWkUve8Mfr64nST28JBgZCkQ6BQ/YcX+rJilLcKvT2wPlrxRcYSU4A==" /></div>
      <button
        class="btn btn-sm btn-with-count js-toggler-target"
        aria-label="Star this repository" title="Star kawatan/Apricot"
        data-ga-click="Repository, click star button, action:blob#show; text:Star">
        <span class="octicon octicon-star"></span>
        Star
      </button>
        <a class="social-count js-social-count" href="/kawatan/Apricot/stargazers">
          10
        </a>
</form>  </div>

  </li>

  <li>
          <!-- </textarea> --><!-- '"` --><form accept-charset="UTF-8" action="/kawatan/Apricot/fork" class="btn-with-count" data-form-nonce="283793ba78ae1eabcbd2e2c40a1b786ff2ce6e7e" method="post"><div style="margin:0;padding:0;display:inline"><input name="utf8" type="hidden" value="&#x2713;" /><input name="authenticity_token" type="hidden" value="oEhRqmFQZafJ2wl03sYo2ge9ecGkoerC7At8QEAZtNVWx26kIBzRtDzbz85OPVHEtQmngtxgWoGZwGilHjNsJw==" /></div>
            <button
                type="submit"
                class="btn btn-sm btn-with-count"
                data-ga-click="Repository, show fork modal, action:blob#show; text:Fork"
                title="Fork your own copy of kawatan/Apricot to your account"
                aria-label="Fork your own copy of kawatan/Apricot to your account">
              <span class="octicon octicon-repo-forked"></span>
              Fork
            </button>
</form>
    <a href="/kawatan/Apricot/network" class="social-count">
      0
    </a>
  </li>
</ul>

    <h1 itemscope itemtype="http://data-vocabulary.org/Breadcrumb" class="entry-title public ">
  <span class="octicon octicon-repo"></span>
  <span class="author"><a href="/kawatan" class="url fn" itemprop="url" rel="author"><span itemprop="title">kawatan</span></a></span><!--
--><span class="path-divider">/</span><!--
--><strong><a href="/kawatan/Apricot" data-pjax="#js-repo-pjax-container">Apricot</a></strong>

  <span class="page-context-loader">
    <img alt="" height="16" src="https://assets-cdn.github.com/images/spinners/octocat-spinner-32.gif" width="16" />
  </span>

</h1>

  </div>
  <div class="container">
    
<nav class="reponav js-repo-nav js-sidenav-container-pjax js-octicon-loaders"
     role="navigation"
     data-pjax="#js-repo-pjax-container">

  <a href="/kawatan/Apricot" aria-label="Code" aria-selected="true" class="js-selected-navigation-item selected reponav-item" data-hotkey="g c" data-selected-links="repo_source repo_downloads repo_commits repo_releases repo_tags repo_branches /kawatan/Apricot">
    <span class="octicon octicon-code"></span>
    Code
</a>
    <a href="/kawatan/Apricot/issues" class="js-selected-navigation-item reponav-item" data-hotkey="g i" data-selected-links="repo_issues repo_labels repo_milestones /kawatan/Apricot/issues">
      <span class="octicon octicon-issue-opened"></span>
      Issues
      <span class="counter">0</span>
</a>
  <a href="/kawatan/Apricot/pulls" class="js-selected-navigation-item reponav-item" data-hotkey="g p" data-selected-links="repo_pulls /kawatan/Apricot/pulls">
    <span class="octicon octicon-git-pull-request"></span>
    Pull requests
    <span class="counter">0</span>
</a>
    <a href="/kawatan/Apricot/wiki" class="js-selected-navigation-item reponav-item" data-hotkey="g w" data-selected-links="repo_wiki /kawatan/Apricot/wiki">
      <span class="octicon octicon-book"></span>
      Wiki
</a>
  <a href="/kawatan/Apricot/pulse" class="js-selected-navigation-item reponav-item" data-selected-links="pulse /kawatan/Apricot/pulse">
    <span class="octicon octicon-pulse"></span>
    Pulse
</a>
  <a href="/kawatan/Apricot/graphs" class="js-selected-navigation-item reponav-item" data-selected-links="repo_graphs repo_contributors /kawatan/Apricot/graphs">
    <span class="octicon octicon-graph"></span>
    Graphs
</a>

</nav>

  </div>
</div>

<div class="container new-discussion-timeline experiment-repo-nav">
  <div class="repository-content">

    

<a href="/kawatan/Apricot/blob/84bae666f2d8608e3617c8f0febcbd8c7f3e005c/Weather.py" class="hidden js-permalink-shortcut" data-hotkey="y">Permalink</a>

<!-- blob contrib key: blob_contributors:v21:d9791b6ce2b4145d5b25397e794bf7c9 -->

<div class="file-navigation js-zeroclipboard-container">
  
<div class="select-menu js-menu-container js-select-menu left">
  <button class="btn btn-sm select-menu-button js-menu-target css-truncate" data-hotkey="w"
    title="master"
    type="button" aria-label="Switch branches or tags" tabindex="0" aria-haspopup="true">
    <i>Branch:</i>
    <span class="js-select-button css-truncate-target">master</span>
  </button>

  <div class="select-menu-modal-holder js-menu-content js-navigation-container" data-pjax aria-hidden="true">

    <div class="select-menu-modal">
      <div class="select-menu-header">
        <span class="octicon octicon-x js-menu-close" role="button" aria-label="Close"></span>
        <span class="select-menu-title">Switch branches/tags</span>
      </div>

      <div class="select-menu-filters">
        <div class="select-menu-text-filter">
          <input type="text" aria-label="Filter branches/tags" id="context-commitish-filter-field" class="js-filterable-field js-navigation-enable" placeholder="Filter branches/tags">
        </div>
        <div class="select-menu-tabs">
          <ul>
            <li class="select-menu-tab">
              <a href="#" data-tab-filter="branches" data-filter-placeholder="Filter branches/tags" class="js-select-menu-tab" role="tab">Branches</a>
            </li>
            <li class="select-menu-tab">
              <a href="#" data-tab-filter="tags" data-filter-placeholder="Find a tag…" class="js-select-menu-tab" role="tab">Tags</a>
            </li>
          </ul>
        </div>
      </div>

      <div class="select-menu-list select-menu-tab-bucket js-select-menu-tab-bucket" data-tab-filter="branches" role="menu">

        <div data-filterable-for="context-commitish-filter-field" data-filterable-type="substring">


            <a class="select-menu-item js-navigation-item js-navigation-open "
               href="/kawatan/Apricot/blob/gh-pages/Weather.py"
               data-name="gh-pages"
               data-skip-pjax="true"
               rel="nofollow">
              <span class="select-menu-item-icon octicon octicon-check"></span>
              <span class="select-menu-item-text css-truncate-target" title="gh-pages">
                gh-pages
              </span>
            </a>
            <a class="select-menu-item js-navigation-item js-navigation-open selected"
               href="/kawatan/Apricot/blob/master/Weather.py"
               data-name="master"
               data-skip-pjax="true"
               rel="nofollow">
              <span class="select-menu-item-icon octicon octicon-check"></span>
              <span class="select-menu-item-text css-truncate-target" title="master">
                master
              </span>
            </a>
        </div>

          <div class="select-menu-no-results">Nothing to show</div>
      </div>

      <div class="select-menu-list select-menu-tab-bucket js-select-menu-tab-bucket" data-tab-filter="tags">
        <div data-filterable-for="context-commitish-filter-field" data-filterable-type="substring">


        </div>

        <div class="select-menu-no-results">Nothing to show</div>
      </div>

    </div>
  </div>
</div>

  <div class="btn-group right">
    <a href="/kawatan/Apricot/find/master"
          class="js-show-file-finder btn btn-sm"
          data-pjax
          data-hotkey="t">
      Find file
    </a>
    <button aria-label="Copy file path to clipboard" class="js-zeroclipboard btn btn-sm zeroclipboard-button tooltipped tooltipped-s" data-copied-hint="Copied!" type="button">Copy path</button>
  </div>
  <div class="breadcrumb js-zeroclipboard-target">
    <span class="repo-root js-repo-root"><span itemscope="" itemtype="http://data-vocabulary.org/Breadcrumb"><a href="/kawatan/Apricot" class="" data-branch="master" data-pjax="true" itemscope="url"><span itemprop="title">Apricot</span></a></span></span><span class="separator">/</span><strong class="final-path">Weather.py</strong>
  </div>
</div>


  <div class="commit-tease">
      <span class="right">
        <a class="commit-tease-sha" href="/kawatan/Apricot/commit/b11cdbcf2961bc07215cda6f735cc5efda8cd369" data-pjax>
          b11cdbc
        </a>
        <time datetime="2014-10-15T16:06:37Z" is="relative-time">Oct 16, 2014</time>
      </span>
      <div>
        <img alt="@kawatan" class="avatar" height="20" src="https://avatars0.githubusercontent.com/u/246691?v=3&amp;s=40" width="20" />
        <a href="/kawatan" class="user-mention" rel="author">kawatan</a>
          <a href="/kawatan/Apricot/commit/b11cdbcf2961bc07215cda6f735cc5efda8cd369" class="message" data-pjax="true" title="Update Weather.py">Update Weather.py</a>
      </div>

    <div class="commit-tease-contributors">
      <a class="muted-link contributors-toggle" href="#blob_contributors_box" rel="facebox">
        <strong>1</strong>
         contributor
      </a>
      
    </div>

    <div id="blob_contributors_box" style="display:none">
      <h2 class="facebox-header" data-facebox-id="facebox-header">Users who have contributed to this file</h2>
      <ul class="facebox-user-list" data-facebox-id="facebox-description">
          <li class="facebox-user-list-item">
            <img alt="@kawatan" height="24" src="https://avatars2.githubusercontent.com/u/246691?v=3&amp;s=48" width="24" />
            <a href="/kawatan">kawatan</a>
          </li>
      </ul>
    </div>
  </div>

<div class="file">
  <div class="file-header">
  <div class="file-actions">

    <div class="btn-group">
      <a href="/kawatan/Apricot/raw/master/Weather.py" class="btn btn-sm " id="raw-url">Raw</a>
        <a href="/kawatan/Apricot/blame/master/Weather.py" class="btn btn-sm js-update-url-with-hash">Blame</a>
      <a href="/kawatan/Apricot/commits/master/Weather.py" class="btn btn-sm " rel="nofollow">History</a>
    </div>

        <a class="octicon-btn tooltipped tooltipped-nw"
           href="github-windows://openRepo/https://github.com/kawatan/Apricot?branch=master&amp;filepath=Weather.py"
           aria-label="Open this file in GitHub Desktop"
           data-ga-click="Repository, open with desktop, type:windows">
            <span class="octicon octicon-device-desktop"></span>
        </a>

        <!-- </textarea> --><!-- '"` --><form accept-charset="UTF-8" action="/kawatan/Apricot/edit/master/Weather.py" class="inline-form js-update-url-with-hash" data-form-nonce="283793ba78ae1eabcbd2e2c40a1b786ff2ce6e7e" method="post"><div style="margin:0;padding:0;display:inline"><input name="utf8" type="hidden" value="&#x2713;" /><input name="authenticity_token" type="hidden" value="MrG34ksI0Lw/iZCXYzUNgpata/fLkDcEtoqXYjTu0OOqseh100Q4bOmYLMkEO54c2tzpvuwkqhMK5huIFoPguw==" /></div>
          <button class="octicon-btn tooltipped tooltipped-nw" type="submit"
            aria-label="Fork this project and edit the file" data-hotkey="e" data-disable-with>
            <span class="octicon octicon-pencil"></span>
          </button>
</form>        <!-- </textarea> --><!-- '"` --><form accept-charset="UTF-8" action="/kawatan/Apricot/delete/master/Weather.py" class="inline-form" data-form-nonce="283793ba78ae1eabcbd2e2c40a1b786ff2ce6e7e" method="post"><div style="margin:0;padding:0;display:inline"><input name="utf8" type="hidden" value="&#x2713;" /><input name="authenticity_token" type="hidden" value="pPPSIXctEXy/f2r0ODbgZCJJqfnG8eF9n9YV96OnXUdlCm3/FvvXJpnZXlDE+ZpRRPrtBWslUp93K4kHTiIVCA==" /></div>
          <button class="octicon-btn octicon-btn-danger tooltipped tooltipped-nw" type="submit"
            aria-label="Fork this project and delete the file" data-disable-with>
            <span class="octicon octicon-trashcan"></span>
          </button>
</form>  </div>

  <div class="file-info">
      <span class="file-mode" title="File mode">executable file</span>
      <span class="file-info-divider"></span>
      842 lines (632 sloc)
      <span class="file-info-divider"></span>
    29.1 KB
  </div>
</div>

  

  <div class="blob-wrapper data type-python">
      <table class="highlight tab-size js-file-line-container" data-tab-size="8">
      <tr>
        <td id="L1" class="blob-num js-line-number" data-line-number="1"></td>
        <td id="LC1" class="blob-code blob-code-inner js-file-line"><span class="pl-c"># -*- coding: utf-8 -*-</span></td>
      </tr>
      <tr>
        <td id="L2" class="blob-num js-line-number" data-line-number="2"></td>
        <td id="LC2" class="blob-code blob-code-inner js-file-line"><span class="pl-c"># Weather.py</span></td>
      </tr>
      <tr>
        <td id="L3" class="blob-num js-line-number" data-line-number="3"></td>
        <td id="LC3" class="blob-code blob-code-inner js-file-line"><span class="pl-c"># Copyright © Masaaki Kawata All rights reserved.</span></td>
      </tr>
      <tr>
        <td id="L4" class="blob-num js-line-number" data-line-number="4"></td>
        <td id="LC4" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L5" class="blob-num js-line-number" data-line-number="5"></td>
        <td id="LC5" class="blob-code blob-code-inner js-file-line"><span class="pl-k">import</span> clr</td>
      </tr>
      <tr>
        <td id="L6" class="blob-num js-line-number" data-line-number="6"></td>
        <td id="LC6" class="blob-code blob-code-inner js-file-line">clr.AddReferenceByPartialName(<span class="pl-s"><span class="pl-pds">&quot;</span>mscorlib<span class="pl-pds">&quot;</span></span>)</td>
      </tr>
      <tr>
        <td id="L7" class="blob-num js-line-number" data-line-number="7"></td>
        <td id="LC7" class="blob-code blob-code-inner js-file-line">clr.AddReferenceByPartialName(<span class="pl-s"><span class="pl-pds">&quot;</span>System<span class="pl-pds">&quot;</span></span>)</td>
      </tr>
      <tr>
        <td id="L8" class="blob-num js-line-number" data-line-number="8"></td>
        <td id="LC8" class="blob-code blob-code-inner js-file-line">clr.AddReferenceByPartialName(<span class="pl-s"><span class="pl-pds">&quot;</span>System.Core<span class="pl-pds">&quot;</span></span>)</td>
      </tr>
      <tr>
        <td id="L9" class="blob-num js-line-number" data-line-number="9"></td>
        <td id="LC9" class="blob-code blob-code-inner js-file-line">clr.AddReferenceByPartialName(<span class="pl-s"><span class="pl-pds">&quot;</span>WindowsBase<span class="pl-pds">&quot;</span></span>)</td>
      </tr>
      <tr>
        <td id="L10" class="blob-num js-line-number" data-line-number="10"></td>
        <td id="LC10" class="blob-code blob-code-inner js-file-line">clr.AddReferenceByPartialName(<span class="pl-s"><span class="pl-pds">&quot;</span>PresentationCore<span class="pl-pds">&quot;</span></span>)</td>
      </tr>
      <tr>
        <td id="L11" class="blob-num js-line-number" data-line-number="11"></td>
        <td id="LC11" class="blob-code blob-code-inner js-file-line">clr.AddReferenceByPartialName(<span class="pl-s"><span class="pl-pds">&quot;</span>PresentationFramework<span class="pl-pds">&quot;</span></span>)</td>
      </tr>
      <tr>
        <td id="L12" class="blob-num js-line-number" data-line-number="12"></td>
        <td id="LC12" class="blob-code blob-code-inner js-file-line">clr.AddReferenceByPartialName(<span class="pl-s"><span class="pl-pds">&quot;</span>Apricot<span class="pl-pds">&quot;</span></span>)</td>
      </tr>
      <tr>
        <td id="L13" class="blob-num js-line-number" data-line-number="13"></td>
        <td id="LC13" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L14" class="blob-num js-line-number" data-line-number="14"></td>
        <td id="LC14" class="blob-code blob-code-inner js-file-line"><span class="pl-k">from</span> System <span class="pl-k">import</span> Object, Nullable, Byte, UInt32, Double, Char, String, Uri, DateTime, TimeSpan, Array, StringComparison, Convert, BitConverter, Math, Action</td>
      </tr>
      <tr>
        <td id="L15" class="blob-num js-line-number" data-line-number="15"></td>
        <td id="LC15" class="blob-code blob-code-inner js-file-line"><span class="pl-k">from</span> System.Collections.Generic <span class="pl-k">import</span> List, Dictionary, KeyValuePair, HashSet</td>
      </tr>
      <tr>
        <td id="L16" class="blob-num js-line-number" data-line-number="16"></td>
        <td id="LC16" class="blob-code blob-code-inner js-file-line"><span class="pl-k">from</span> System.IO <span class="pl-k">import</span> Stream, StreamReader, MemoryStream, FileStream, SeekOrigin, FileMode, FileAccess, FileShare</td>
      </tr>
      <tr>
        <td id="L17" class="blob-num js-line-number" data-line-number="17"></td>
        <td id="LC17" class="blob-code blob-code-inner js-file-line"><span class="pl-k">from</span> System.Diagnostics <span class="pl-k">import</span> Trace</td>
      </tr>
      <tr>
        <td id="L18" class="blob-num js-line-number" data-line-number="18"></td>
        <td id="LC18" class="blob-code blob-code-inner js-file-line"><span class="pl-k">from</span> System.Globalization <span class="pl-k">import</span> CultureInfo, NumberStyles</td>
      </tr>
      <tr>
        <td id="L19" class="blob-num js-line-number" data-line-number="19"></td>
        <td id="LC19" class="blob-code blob-code-inner js-file-line"><span class="pl-k">from</span> System.Text <span class="pl-k">import</span> StringBuilder, Encoding</td>
      </tr>
      <tr>
        <td id="L20" class="blob-num js-line-number" data-line-number="20"></td>
        <td id="LC20" class="blob-code blob-code-inner js-file-line"><span class="pl-k">from</span> System.Text.RegularExpressions <span class="pl-k">import</span> Regex, Match, RegexOptions</td>
      </tr>
      <tr>
        <td id="L21" class="blob-num js-line-number" data-line-number="21"></td>
        <td id="LC21" class="blob-code blob-code-inner js-file-line"><span class="pl-k">from</span> System.Threading.Tasks <span class="pl-k">import</span> Task, TaskCreationOptions, TaskScheduler</td>
      </tr>
      <tr>
        <td id="L22" class="blob-num js-line-number" data-line-number="22"></td>
        <td id="LC22" class="blob-code blob-code-inner js-file-line"><span class="pl-k">from</span> System.Linq <span class="pl-k">import</span> Enumerable</td>
      </tr>
      <tr>
        <td id="L23" class="blob-num js-line-number" data-line-number="23"></td>
        <td id="LC23" class="blob-code blob-code-inner js-file-line"><span class="pl-k">from</span> System.Net <span class="pl-k">import</span> WebRequest, WebResponse</td>
      </tr>
      <tr>
        <td id="L24" class="blob-num js-line-number" data-line-number="24"></td>
        <td id="LC24" class="blob-code blob-code-inner js-file-line"><span class="pl-k">from</span> System.Net.NetworkInformation <span class="pl-k">import</span> NetworkInterface</td>
      </tr>
      <tr>
        <td id="L25" class="blob-num js-line-number" data-line-number="25"></td>
        <td id="LC25" class="blob-code blob-code-inner js-file-line"><span class="pl-k">from</span> System.Windows <span class="pl-k">import</span> Application, Window, WindowStartupLocation, WindowStyle, ResizeMode, SizeToContent, HorizontalAlignment, VerticalAlignment, Rect, PropertyPath</td>
      </tr>
      <tr>
        <td id="L26" class="blob-num js-line-number" data-line-number="26"></td>
        <td id="LC26" class="blob-code blob-code-inner js-file-line"><span class="pl-k">from</span> System.Windows.Controls <span class="pl-k">import</span> ContentControl, Grid</td>
      </tr>
      <tr>
        <td id="L27" class="blob-num js-line-number" data-line-number="27"></td>
        <td id="LC27" class="blob-code blob-code-inner js-file-line"><span class="pl-k">from</span> System.Windows.Media <span class="pl-k">import</span> Color, Colors, Brushes, SolidColorBrush, TranslateTransform, ScaleTransform, RectangleGeometry, EllipseGeometry, ImageBrush, TileMode, BrushMappingMode, Stretch, AlignmentX, AlignmentY, DrawingGroup, DrawingContext, DrawingImage</td>
      </tr>
      <tr>
        <td id="L28" class="blob-num js-line-number" data-line-number="28"></td>
        <td id="LC28" class="blob-code blob-code-inner js-file-line"><span class="pl-k">from</span> System.Windows.Media.Animation <span class="pl-k">import</span> Storyboard, Clock, ClockState, DoubleAnimation, SineEase, EasingMode</td>
      </tr>
      <tr>
        <td id="L29" class="blob-num js-line-number" data-line-number="29"></td>
        <td id="LC29" class="blob-code blob-code-inner js-file-line"><span class="pl-k">from</span> System.Windows.Media.Effects <span class="pl-k">import</span> DropShadowEffect</td>
      </tr>
      <tr>
        <td id="L30" class="blob-num js-line-number" data-line-number="30"></td>
        <td id="LC30" class="blob-code blob-code-inner js-file-line"><span class="pl-k">from</span> System.Windows.Media.Imaging <span class="pl-k">import</span> BitmapImage, BitmapCacheOption, BitmapCreateOptions</td>
      </tr>
      <tr>
        <td id="L31" class="blob-num js-line-number" data-line-number="31"></td>
        <td id="LC31" class="blob-code blob-code-inner js-file-line"><span class="pl-k">from</span> System.Windows.Shapes <span class="pl-k">import</span> Rectangle</td>
      </tr>
      <tr>
        <td id="L32" class="blob-num js-line-number" data-line-number="32"></td>
        <td id="LC32" class="blob-code blob-code-inner js-file-line"><span class="pl-k">from</span> System.Windows.Threading <span class="pl-k">import</span> DispatcherTimer, DispatcherPriority</td>
      </tr>
      <tr>
        <td id="L33" class="blob-num js-line-number" data-line-number="33"></td>
        <td id="LC33" class="blob-code blob-code-inner js-file-line"><span class="pl-k">from</span> Apricot <span class="pl-k">import</span> Script, Sequence</td>
      </tr>
      <tr>
        <td id="L34" class="blob-num js-line-number" data-line-number="34"></td>
        <td id="LC34" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L35" class="blob-num js-line-number" data-line-number="35"></td>
        <td id="LC35" class="blob-code blob-code-inner js-file-line"><span class="pl-c"># e.g. Tokyo, Japan</span></td>
      </tr>
      <tr>
        <td id="L36" class="blob-num js-line-number" data-line-number="36"></td>
        <td id="LC36" class="blob-code blob-code-inner js-file-line">location <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">&quot;</span>Tokyo, Japan<span class="pl-pds">&quot;</span></span></td>
      </tr>
      <tr>
        <td id="L37" class="blob-num js-line-number" data-line-number="37"></td>
        <td id="LC37" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L38" class="blob-num js-line-number" data-line-number="38"></td>
        <td id="LC38" class="blob-code blob-code-inner js-file-line"><span class="pl-c"># http://www.json.org/</span></td>
      </tr>
      <tr>
        <td id="L39" class="blob-num js-line-number" data-line-number="39"></td>
        <td id="LC39" class="blob-code blob-code-inner js-file-line"><span class="pl-k">class</span> <span class="pl-en">JsonDecoder</span>(<span class="pl-e">Object</span>):</td>
      </tr>
      <tr>
        <td id="L40" class="blob-num js-line-number" data-line-number="40"></td>
        <td id="LC40" class="blob-code blob-code-inner js-file-line">	TOKEN_NONE <span class="pl-k">=</span> <span class="pl-c1">0</span></td>
      </tr>
      <tr>
        <td id="L41" class="blob-num js-line-number" data-line-number="41"></td>
        <td id="LC41" class="blob-code blob-code-inner js-file-line">	TOKEN_CURLY_OPEN <span class="pl-k">=</span> <span class="pl-c1">1</span></td>
      </tr>
      <tr>
        <td id="L42" class="blob-num js-line-number" data-line-number="42"></td>
        <td id="LC42" class="blob-code blob-code-inner js-file-line">	TOKEN_CURLY_CLOSE <span class="pl-k">=</span> <span class="pl-c1">2</span></td>
      </tr>
      <tr>
        <td id="L43" class="blob-num js-line-number" data-line-number="43"></td>
        <td id="LC43" class="blob-code blob-code-inner js-file-line">	TOKEN_SQUARED_OPEN <span class="pl-k">=</span> <span class="pl-c1">3</span></td>
      </tr>
      <tr>
        <td id="L44" class="blob-num js-line-number" data-line-number="44"></td>
        <td id="LC44" class="blob-code blob-code-inner js-file-line">	TOKEN_SQUARED_CLOSE <span class="pl-k">=</span> <span class="pl-c1">4</span></td>
      </tr>
      <tr>
        <td id="L45" class="blob-num js-line-number" data-line-number="45"></td>
        <td id="LC45" class="blob-code blob-code-inner js-file-line">	TOKEN_COLON <span class="pl-k">=</span> <span class="pl-c1">5</span></td>
      </tr>
      <tr>
        <td id="L46" class="blob-num js-line-number" data-line-number="46"></td>
        <td id="LC46" class="blob-code blob-code-inner js-file-line">	TOKEN_COMMA <span class="pl-k">=</span> <span class="pl-c1">6</span></td>
      </tr>
      <tr>
        <td id="L47" class="blob-num js-line-number" data-line-number="47"></td>
        <td id="LC47" class="blob-code blob-code-inner js-file-line">	TOKEN_STRING <span class="pl-k">=</span> <span class="pl-c1">7</span></td>
      </tr>
      <tr>
        <td id="L48" class="blob-num js-line-number" data-line-number="48"></td>
        <td id="LC48" class="blob-code blob-code-inner js-file-line">	TOKEN_NUMBER <span class="pl-k">=</span> <span class="pl-c1">8</span></td>
      </tr>
      <tr>
        <td id="L49" class="blob-num js-line-number" data-line-number="49"></td>
        <td id="LC49" class="blob-code blob-code-inner js-file-line">	TOKEN_TRUE <span class="pl-k">=</span> <span class="pl-c1">9</span></td>
      </tr>
      <tr>
        <td id="L50" class="blob-num js-line-number" data-line-number="50"></td>
        <td id="LC50" class="blob-code blob-code-inner js-file-line">	TOKEN_FALSE <span class="pl-k">=</span> <span class="pl-c1">10</span></td>
      </tr>
      <tr>
        <td id="L51" class="blob-num js-line-number" data-line-number="51"></td>
        <td id="LC51" class="blob-code blob-code-inner js-file-line">	TOKEN_NULL <span class="pl-k">=</span> <span class="pl-c1">11</span></td>
      </tr>
      <tr>
        <td id="L52" class="blob-num js-line-number" data-line-number="52"></td>
        <td id="LC52" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L53" class="blob-num js-line-number" data-line-number="53"></td>
        <td id="LC53" class="blob-code blob-code-inner js-file-line">	<span class="pl-en">@<span class="pl-c1">staticmethod</span></span></td>
      </tr>
      <tr>
        <td id="L54" class="blob-num js-line-number" data-line-number="54"></td>
        <td id="LC54" class="blob-code blob-code-inner js-file-line">	<span class="pl-k">def</span> <span class="pl-en">decode</span>(<span class="pl-smi">json</span>):</td>
      </tr>
      <tr>
        <td id="L55" class="blob-num js-line-number" data-line-number="55"></td>
        <td id="LC55" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">if</span> json <span class="pl-k">is</span> <span class="pl-k">not</span> <span class="pl-c1">None</span>:</td>
      </tr>
      <tr>
        <td id="L56" class="blob-num js-line-number" data-line-number="56"></td>
        <td id="LC56" class="blob-code blob-code-inner js-file-line">			charArray <span class="pl-k">=</span> json.ToCharArray()</td>
      </tr>
      <tr>
        <td id="L57" class="blob-num js-line-number" data-line-number="57"></td>
        <td id="LC57" class="blob-code blob-code-inner js-file-line">			index <span class="pl-k">=</span> <span class="pl-c1">0</span></td>
      </tr>
      <tr>
        <td id="L58" class="blob-num js-line-number" data-line-number="58"></td>
        <td id="LC58" class="blob-code blob-code-inner js-file-line">			value, index, success <span class="pl-k">=</span> JsonDecoder.parseValue(charArray, index, <span class="pl-c1">True</span>)</td>
      </tr>
      <tr>
        <td id="L59" class="blob-num js-line-number" data-line-number="59"></td>
        <td id="LC59" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L60" class="blob-num js-line-number" data-line-number="60"></td>
        <td id="LC60" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">return</span> value</td>
      </tr>
      <tr>
        <td id="L61" class="blob-num js-line-number" data-line-number="61"></td>
        <td id="LC61" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L62" class="blob-num js-line-number" data-line-number="62"></td>
        <td id="LC62" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">return</span> <span class="pl-c1">None</span></td>
      </tr>
      <tr>
        <td id="L63" class="blob-num js-line-number" data-line-number="63"></td>
        <td id="LC63" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L64" class="blob-num js-line-number" data-line-number="64"></td>
        <td id="LC64" class="blob-code blob-code-inner js-file-line">	<span class="pl-en">@<span class="pl-c1">staticmethod</span></span></td>
      </tr>
      <tr>
        <td id="L65" class="blob-num js-line-number" data-line-number="65"></td>
        <td id="LC65" class="blob-code blob-code-inner js-file-line">	<span class="pl-k">def</span> <span class="pl-en">parseObject</span>(<span class="pl-smi">json</span>, <span class="pl-smi">index</span>, <span class="pl-smi">success</span>):</td>
      </tr>
      <tr>
        <td id="L66" class="blob-num js-line-number" data-line-number="66"></td>
        <td id="LC66" class="blob-code blob-code-inner js-file-line">		dictionary <span class="pl-k">=</span> Dictionary[String, Object]()</td>
      </tr>
      <tr>
        <td id="L67" class="blob-num js-line-number" data-line-number="67"></td>
        <td id="LC67" class="blob-code blob-code-inner js-file-line">		token <span class="pl-k">=</span> <span class="pl-k">-</span><span class="pl-c1">1</span></td>
      </tr>
      <tr>
        <td id="L68" class="blob-num js-line-number" data-line-number="68"></td>
        <td id="LC68" class="blob-code blob-code-inner js-file-line">		i, index <span class="pl-k">=</span> JsonDecoder.nextToken(json, index) <span class="pl-c"># {</span></td>
      </tr>
      <tr>
        <td id="L69" class="blob-num js-line-number" data-line-number="69"></td>
        <td id="LC69" class="blob-code blob-code-inner js-file-line">		done <span class="pl-k">=</span> <span class="pl-c1">False</span></td>
      </tr>
      <tr>
        <td id="L70" class="blob-num js-line-number" data-line-number="70"></td>
        <td id="LC70" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L71" class="blob-num js-line-number" data-line-number="71"></td>
        <td id="LC71" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">while</span> <span class="pl-k">not</span> done:</td>
      </tr>
      <tr>
        <td id="L72" class="blob-num js-line-number" data-line-number="72"></td>
        <td id="LC72" class="blob-code blob-code-inner js-file-line">			token <span class="pl-k">=</span> JsonDecoder.lookAhead(json, index)</td>
      </tr>
      <tr>
        <td id="L73" class="blob-num js-line-number" data-line-number="73"></td>
        <td id="LC73" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L74" class="blob-num js-line-number" data-line-number="74"></td>
        <td id="LC74" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">if</span> token <span class="pl-k">==</span> JsonDecoder.TOKEN_NONE:</td>
      </tr>
      <tr>
        <td id="L75" class="blob-num js-line-number" data-line-number="75"></td>
        <td id="LC75" class="blob-code blob-code-inner js-file-line">				<span class="pl-k">return</span> <span class="pl-c1">None</span>, index, <span class="pl-c1">False</span></td>
      </tr>
      <tr>
        <td id="L76" class="blob-num js-line-number" data-line-number="76"></td>
        <td id="LC76" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L77" class="blob-num js-line-number" data-line-number="77"></td>
        <td id="LC77" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">elif</span> token <span class="pl-k">==</span> JsonDecoder.TOKEN_COMMA:</td>
      </tr>
      <tr>
        <td id="L78" class="blob-num js-line-number" data-line-number="78"></td>
        <td id="LC78" class="blob-code blob-code-inner js-file-line">				i, index <span class="pl-k">=</span> JsonDecoder.nextToken(json, index)</td>
      </tr>
      <tr>
        <td id="L79" class="blob-num js-line-number" data-line-number="79"></td>
        <td id="LC79" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L80" class="blob-num js-line-number" data-line-number="80"></td>
        <td id="LC80" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">elif</span> token <span class="pl-k">==</span> JsonDecoder.TOKEN_CURLY_CLOSE:</td>
      </tr>
      <tr>
        <td id="L81" class="blob-num js-line-number" data-line-number="81"></td>
        <td id="LC81" class="blob-code blob-code-inner js-file-line">				i, index <span class="pl-k">=</span> JsonDecoder.nextToken(json, index)</td>
      </tr>
      <tr>
        <td id="L82" class="blob-num js-line-number" data-line-number="82"></td>
        <td id="LC82" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L83" class="blob-num js-line-number" data-line-number="83"></td>
        <td id="LC83" class="blob-code blob-code-inner js-file-line">				<span class="pl-k">return</span> dictionary, index, success</td>
      </tr>
      <tr>
        <td id="L84" class="blob-num js-line-number" data-line-number="84"></td>
        <td id="LC84" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L85" class="blob-num js-line-number" data-line-number="85"></td>
        <td id="LC85" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">else</span>:</td>
      </tr>
      <tr>
        <td id="L86" class="blob-num js-line-number" data-line-number="86"></td>
        <td id="LC86" class="blob-code blob-code-inner js-file-line">				name, index, success <span class="pl-k">=</span> JsonDecoder.parseString(json, index, success)</td>
      </tr>
      <tr>
        <td id="L87" class="blob-num js-line-number" data-line-number="87"></td>
        <td id="LC87" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L88" class="blob-num js-line-number" data-line-number="88"></td>
        <td id="LC88" class="blob-code blob-code-inner js-file-line">				<span class="pl-k">if</span> <span class="pl-k">not</span> success:</td>
      </tr>
      <tr>
        <td id="L89" class="blob-num js-line-number" data-line-number="89"></td>
        <td id="LC89" class="blob-code blob-code-inner js-file-line">					<span class="pl-k">return</span> <span class="pl-c1">None</span>, index, success</td>
      </tr>
      <tr>
        <td id="L90" class="blob-num js-line-number" data-line-number="90"></td>
        <td id="LC90" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L91" class="blob-num js-line-number" data-line-number="91"></td>
        <td id="LC91" class="blob-code blob-code-inner js-file-line">				token, index <span class="pl-k">=</span> JsonDecoder.nextToken(json, index)</td>
      </tr>
      <tr>
        <td id="L92" class="blob-num js-line-number" data-line-number="92"></td>
        <td id="LC92" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L93" class="blob-num js-line-number" data-line-number="93"></td>
        <td id="LC93" class="blob-code blob-code-inner js-file-line">				<span class="pl-k">if</span> token <span class="pl-k">!=</span> JsonDecoder.TOKEN_COLON:</td>
      </tr>
      <tr>
        <td id="L94" class="blob-num js-line-number" data-line-number="94"></td>
        <td id="LC94" class="blob-code blob-code-inner js-file-line">					<span class="pl-k">return</span> <span class="pl-c1">None</span>, index, <span class="pl-c1">False</span></td>
      </tr>
      <tr>
        <td id="L95" class="blob-num js-line-number" data-line-number="95"></td>
        <td id="LC95" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L96" class="blob-num js-line-number" data-line-number="96"></td>
        <td id="LC96" class="blob-code blob-code-inner js-file-line">				value, index, success <span class="pl-k">=</span> JsonDecoder.parseValue(json, index, success)</td>
      </tr>
      <tr>
        <td id="L97" class="blob-num js-line-number" data-line-number="97"></td>
        <td id="LC97" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L98" class="blob-num js-line-number" data-line-number="98"></td>
        <td id="LC98" class="blob-code blob-code-inner js-file-line">				<span class="pl-k">if</span> <span class="pl-k">not</span> success:</td>
      </tr>
      <tr>
        <td id="L99" class="blob-num js-line-number" data-line-number="99"></td>
        <td id="LC99" class="blob-code blob-code-inner js-file-line">					<span class="pl-k">return</span> <span class="pl-c1">None</span>, index, success</td>
      </tr>
      <tr>
        <td id="L100" class="blob-num js-line-number" data-line-number="100"></td>
        <td id="LC100" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L101" class="blob-num js-line-number" data-line-number="101"></td>
        <td id="LC101" class="blob-code blob-code-inner js-file-line">				dictionary.Add(name, value)</td>
      </tr>
      <tr>
        <td id="L102" class="blob-num js-line-number" data-line-number="102"></td>
        <td id="LC102" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L103" class="blob-num js-line-number" data-line-number="103"></td>
        <td id="LC103" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">return</span> dictionary, index, success</td>
      </tr>
      <tr>
        <td id="L104" class="blob-num js-line-number" data-line-number="104"></td>
        <td id="LC104" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L105" class="blob-num js-line-number" data-line-number="105"></td>
        <td id="LC105" class="blob-code blob-code-inner js-file-line">	<span class="pl-en">@<span class="pl-c1">staticmethod</span></span></td>
      </tr>
      <tr>
        <td id="L106" class="blob-num js-line-number" data-line-number="106"></td>
        <td id="LC106" class="blob-code blob-code-inner js-file-line">	<span class="pl-k">def</span> <span class="pl-en">parseArray</span>(<span class="pl-smi">json</span>, <span class="pl-smi">index</span>, <span class="pl-smi">success</span>):</td>
      </tr>
      <tr>
        <td id="L107" class="blob-num js-line-number" data-line-number="107"></td>
        <td id="LC107" class="blob-code blob-code-inner js-file-line">		<span class="pl-c1">list</span> <span class="pl-k">=</span> List[Object]()</td>
      </tr>
      <tr>
        <td id="L108" class="blob-num js-line-number" data-line-number="108"></td>
        <td id="LC108" class="blob-code blob-code-inner js-file-line">		i, index <span class="pl-k">=</span> JsonDecoder.nextToken(json, index) <span class="pl-c"># [</span></td>
      </tr>
      <tr>
        <td id="L109" class="blob-num js-line-number" data-line-number="109"></td>
        <td id="LC109" class="blob-code blob-code-inner js-file-line">		done <span class="pl-k">=</span> <span class="pl-c1">False</span></td>
      </tr>
      <tr>
        <td id="L110" class="blob-num js-line-number" data-line-number="110"></td>
        <td id="LC110" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L111" class="blob-num js-line-number" data-line-number="111"></td>
        <td id="LC111" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">while</span> <span class="pl-k">not</span> done:</td>
      </tr>
      <tr>
        <td id="L112" class="blob-num js-line-number" data-line-number="112"></td>
        <td id="LC112" class="blob-code blob-code-inner js-file-line">			token <span class="pl-k">=</span> JsonDecoder.lookAhead(json, index)</td>
      </tr>
      <tr>
        <td id="L113" class="blob-num js-line-number" data-line-number="113"></td>
        <td id="LC113" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L114" class="blob-num js-line-number" data-line-number="114"></td>
        <td id="LC114" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">if</span> token <span class="pl-k">==</span> JsonDecoder.TOKEN_NONE:</td>
      </tr>
      <tr>
        <td id="L115" class="blob-num js-line-number" data-line-number="115"></td>
        <td id="LC115" class="blob-code blob-code-inner js-file-line">				<span class="pl-k">return</span> <span class="pl-c1">None</span>, index, <span class="pl-c1">False</span></td>
      </tr>
      <tr>
        <td id="L116" class="blob-num js-line-number" data-line-number="116"></td>
        <td id="LC116" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L117" class="blob-num js-line-number" data-line-number="117"></td>
        <td id="LC117" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">elif</span> token <span class="pl-k">==</span> JsonDecoder.TOKEN_COMMA:</td>
      </tr>
      <tr>
        <td id="L118" class="blob-num js-line-number" data-line-number="118"></td>
        <td id="LC118" class="blob-code blob-code-inner js-file-line">				i, index <span class="pl-k">=</span> JsonDecoder.nextToken(json, index)</td>
      </tr>
      <tr>
        <td id="L119" class="blob-num js-line-number" data-line-number="119"></td>
        <td id="LC119" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L120" class="blob-num js-line-number" data-line-number="120"></td>
        <td id="LC120" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">elif</span> token <span class="pl-k">==</span> JsonDecoder.TOKEN_SQUARED_CLOSE:</td>
      </tr>
      <tr>
        <td id="L121" class="blob-num js-line-number" data-line-number="121"></td>
        <td id="LC121" class="blob-code blob-code-inner js-file-line">				i, index <span class="pl-k">=</span> JsonDecoder.nextToken(json, index)</td>
      </tr>
      <tr>
        <td id="L122" class="blob-num js-line-number" data-line-number="122"></td>
        <td id="LC122" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L123" class="blob-num js-line-number" data-line-number="123"></td>
        <td id="LC123" class="blob-code blob-code-inner js-file-line">				<span class="pl-k">break</span></td>
      </tr>
      <tr>
        <td id="L124" class="blob-num js-line-number" data-line-number="124"></td>
        <td id="LC124" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L125" class="blob-num js-line-number" data-line-number="125"></td>
        <td id="LC125" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">else</span>:</td>
      </tr>
      <tr>
        <td id="L126" class="blob-num js-line-number" data-line-number="126"></td>
        <td id="LC126" class="blob-code blob-code-inner js-file-line">				value, index, success <span class="pl-k">=</span> JsonDecoder.parseValue(json, index, success)</td>
      </tr>
      <tr>
        <td id="L127" class="blob-num js-line-number" data-line-number="127"></td>
        <td id="LC127" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L128" class="blob-num js-line-number" data-line-number="128"></td>
        <td id="LC128" class="blob-code blob-code-inner js-file-line">				<span class="pl-k">if</span> <span class="pl-k">not</span> success:</td>
      </tr>
      <tr>
        <td id="L129" class="blob-num js-line-number" data-line-number="129"></td>
        <td id="LC129" class="blob-code blob-code-inner js-file-line">					<span class="pl-k">return</span> <span class="pl-c1">None</span>, index, success</td>
      </tr>
      <tr>
        <td id="L130" class="blob-num js-line-number" data-line-number="130"></td>
        <td id="LC130" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L131" class="blob-num js-line-number" data-line-number="131"></td>
        <td id="LC131" class="blob-code blob-code-inner js-file-line">				<span class="pl-c1">list</span>.Add(value)</td>
      </tr>
      <tr>
        <td id="L132" class="blob-num js-line-number" data-line-number="132"></td>
        <td id="LC132" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L133" class="blob-num js-line-number" data-line-number="133"></td>
        <td id="LC133" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">return</span> <span class="pl-c1">list</span>.ToArray(), index, success</td>
      </tr>
      <tr>
        <td id="L134" class="blob-num js-line-number" data-line-number="134"></td>
        <td id="LC134" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L135" class="blob-num js-line-number" data-line-number="135"></td>
        <td id="LC135" class="blob-code blob-code-inner js-file-line">	<span class="pl-en">@<span class="pl-c1">staticmethod</span></span></td>
      </tr>
      <tr>
        <td id="L136" class="blob-num js-line-number" data-line-number="136"></td>
        <td id="LC136" class="blob-code blob-code-inner js-file-line">	<span class="pl-k">def</span> <span class="pl-en">parseValue</span>(<span class="pl-smi">json</span>, <span class="pl-smi">index</span>, <span class="pl-smi">success</span>):</td>
      </tr>
      <tr>
        <td id="L137" class="blob-num js-line-number" data-line-number="137"></td>
        <td id="LC137" class="blob-code blob-code-inner js-file-line">		i <span class="pl-k">=</span> JsonDecoder.lookAhead(json, index)</td>
      </tr>
      <tr>
        <td id="L138" class="blob-num js-line-number" data-line-number="138"></td>
        <td id="LC138" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L139" class="blob-num js-line-number" data-line-number="139"></td>
        <td id="LC139" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">if</span> i <span class="pl-k">==</span> JsonDecoder.TOKEN_STRING:</td>
      </tr>
      <tr>
        <td id="L140" class="blob-num js-line-number" data-line-number="140"></td>
        <td id="LC140" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">return</span> JsonDecoder.parseString(json, index, success)</td>
      </tr>
      <tr>
        <td id="L141" class="blob-num js-line-number" data-line-number="141"></td>
        <td id="LC141" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L142" class="blob-num js-line-number" data-line-number="142"></td>
        <td id="LC142" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">elif</span> i <span class="pl-k">==</span> JsonDecoder.TOKEN_NUMBER:</td>
      </tr>
      <tr>
        <td id="L143" class="blob-num js-line-number" data-line-number="143"></td>
        <td id="LC143" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">return</span> JsonDecoder.parseNumber(json, index, success)</td>
      </tr>
      <tr>
        <td id="L144" class="blob-num js-line-number" data-line-number="144"></td>
        <td id="LC144" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L145" class="blob-num js-line-number" data-line-number="145"></td>
        <td id="LC145" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">elif</span> i <span class="pl-k">==</span> JsonDecoder.TOKEN_CURLY_OPEN:</td>
      </tr>
      <tr>
        <td id="L146" class="blob-num js-line-number" data-line-number="146"></td>
        <td id="LC146" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">return</span> JsonDecoder.parseObject(json, index, success)</td>
      </tr>
      <tr>
        <td id="L147" class="blob-num js-line-number" data-line-number="147"></td>
        <td id="LC147" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L148" class="blob-num js-line-number" data-line-number="148"></td>
        <td id="LC148" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">elif</span> i <span class="pl-k">==</span> JsonDecoder.TOKEN_SQUARED_OPEN:</td>
      </tr>
      <tr>
        <td id="L149" class="blob-num js-line-number" data-line-number="149"></td>
        <td id="LC149" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">return</span> JsonDecoder.parseArray(json, index, success)</td>
      </tr>
      <tr>
        <td id="L150" class="blob-num js-line-number" data-line-number="150"></td>
        <td id="LC150" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L151" class="blob-num js-line-number" data-line-number="151"></td>
        <td id="LC151" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">elif</span> i <span class="pl-k">==</span> JsonDecoder.TOKEN_TRUE:</td>
      </tr>
      <tr>
        <td id="L152" class="blob-num js-line-number" data-line-number="152"></td>
        <td id="LC152" class="blob-code blob-code-inner js-file-line">			j, index <span class="pl-k">=</span> JsonDecoder.nextToken(json, index)</td>
      </tr>
      <tr>
        <td id="L153" class="blob-num js-line-number" data-line-number="153"></td>
        <td id="LC153" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L154" class="blob-num js-line-number" data-line-number="154"></td>
        <td id="LC154" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">return</span> <span class="pl-c1">True</span>, index, success</td>
      </tr>
      <tr>
        <td id="L155" class="blob-num js-line-number" data-line-number="155"></td>
        <td id="LC155" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L156" class="blob-num js-line-number" data-line-number="156"></td>
        <td id="LC156" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">elif</span> i <span class="pl-k">==</span> JsonDecoder.TOKEN_FALSE:</td>
      </tr>
      <tr>
        <td id="L157" class="blob-num js-line-number" data-line-number="157"></td>
        <td id="LC157" class="blob-code blob-code-inner js-file-line">			j, index <span class="pl-k">=</span> JsonDecoder.nextToken(json, index)</td>
      </tr>
      <tr>
        <td id="L158" class="blob-num js-line-number" data-line-number="158"></td>
        <td id="LC158" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L159" class="blob-num js-line-number" data-line-number="159"></td>
        <td id="LC159" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">return</span> <span class="pl-c1">False</span>, index, success</td>
      </tr>
      <tr>
        <td id="L160" class="blob-num js-line-number" data-line-number="160"></td>
        <td id="LC160" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L161" class="blob-num js-line-number" data-line-number="161"></td>
        <td id="LC161" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">elif</span> i <span class="pl-k">==</span> JsonDecoder.TOKEN_NULL:</td>
      </tr>
      <tr>
        <td id="L162" class="blob-num js-line-number" data-line-number="162"></td>
        <td id="LC162" class="blob-code blob-code-inner js-file-line">			j, index <span class="pl-k">=</span> JsonDecoder.nextToken(json, index)</td>
      </tr>
      <tr>
        <td id="L163" class="blob-num js-line-number" data-line-number="163"></td>
        <td id="LC163" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L164" class="blob-num js-line-number" data-line-number="164"></td>
        <td id="LC164" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">return</span> <span class="pl-c1">None</span>, index, success</td>
      </tr>
      <tr>
        <td id="L165" class="blob-num js-line-number" data-line-number="165"></td>
        <td id="LC165" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L166" class="blob-num js-line-number" data-line-number="166"></td>
        <td id="LC166" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">return</span> <span class="pl-c1">None</span>, index, <span class="pl-c1">False</span></td>
      </tr>
      <tr>
        <td id="L167" class="blob-num js-line-number" data-line-number="167"></td>
        <td id="LC167" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L168" class="blob-num js-line-number" data-line-number="168"></td>
        <td id="LC168" class="blob-code blob-code-inner js-file-line">	<span class="pl-en">@<span class="pl-c1">staticmethod</span></span></td>
      </tr>
      <tr>
        <td id="L169" class="blob-num js-line-number" data-line-number="169"></td>
        <td id="LC169" class="blob-code blob-code-inner js-file-line">	<span class="pl-k">def</span> <span class="pl-en">parseString</span>(<span class="pl-smi">json</span>, <span class="pl-smi">index</span>, <span class="pl-smi">success</span>):</td>
      </tr>
      <tr>
        <td id="L170" class="blob-num js-line-number" data-line-number="170"></td>
        <td id="LC170" class="blob-code blob-code-inner js-file-line">		s <span class="pl-k">=</span> StringBuilder()</td>
      </tr>
      <tr>
        <td id="L171" class="blob-num js-line-number" data-line-number="171"></td>
        <td id="LC171" class="blob-code blob-code-inner js-file-line">		index <span class="pl-k">=</span> JsonDecoder.skipWhitespace(json, index)</td>
      </tr>
      <tr>
        <td id="L172" class="blob-num js-line-number" data-line-number="172"></td>
        <td id="LC172" class="blob-code blob-code-inner js-file-line">		c <span class="pl-k">=</span> json[index] <span class="pl-c"># &quot;</span></td>
      </tr>
      <tr>
        <td id="L173" class="blob-num js-line-number" data-line-number="173"></td>
        <td id="LC173" class="blob-code blob-code-inner js-file-line">		index <span class="pl-k">+=</span> <span class="pl-c1">1</span></td>
      </tr>
      <tr>
        <td id="L174" class="blob-num js-line-number" data-line-number="174"></td>
        <td id="LC174" class="blob-code blob-code-inner js-file-line">		complete <span class="pl-k">=</span> <span class="pl-c1">False</span></td>
      </tr>
      <tr>
        <td id="L175" class="blob-num js-line-number" data-line-number="175"></td>
        <td id="LC175" class="blob-code blob-code-inner js-file-line">		</td>
      </tr>
      <tr>
        <td id="L176" class="blob-num js-line-number" data-line-number="176"></td>
        <td id="LC176" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">while</span> <span class="pl-k">not</span> complete:</td>
      </tr>
      <tr>
        <td id="L177" class="blob-num js-line-number" data-line-number="177"></td>
        <td id="LC177" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">if</span> index <span class="pl-k">==</span> json.Length:</td>
      </tr>
      <tr>
        <td id="L178" class="blob-num js-line-number" data-line-number="178"></td>
        <td id="LC178" class="blob-code blob-code-inner js-file-line">				<span class="pl-k">break</span></td>
      </tr>
      <tr>
        <td id="L179" class="blob-num js-line-number" data-line-number="179"></td>
        <td id="LC179" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L180" class="blob-num js-line-number" data-line-number="180"></td>
        <td id="LC180" class="blob-code blob-code-inner js-file-line">			c <span class="pl-k">=</span> json[index]</td>
      </tr>
      <tr>
        <td id="L181" class="blob-num js-line-number" data-line-number="181"></td>
        <td id="LC181" class="blob-code blob-code-inner js-file-line">			index <span class="pl-k">+=</span> <span class="pl-c1">1</span></td>
      </tr>
      <tr>
        <td id="L182" class="blob-num js-line-number" data-line-number="182"></td>
        <td id="LC182" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L183" class="blob-num js-line-number" data-line-number="183"></td>
        <td id="LC183" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">if</span> c <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>&quot;<span class="pl-pds">&#39;</span></span>:</td>
      </tr>
      <tr>
        <td id="L184" class="blob-num js-line-number" data-line-number="184"></td>
        <td id="LC184" class="blob-code blob-code-inner js-file-line">				complete <span class="pl-k">=</span> <span class="pl-c1">True</span></td>
      </tr>
      <tr>
        <td id="L185" class="blob-num js-line-number" data-line-number="185"></td>
        <td id="LC185" class="blob-code blob-code-inner js-file-line">				<span class="pl-k">break</span></td>
      </tr>
      <tr>
        <td id="L186" class="blob-num js-line-number" data-line-number="186"></td>
        <td id="LC186" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L187" class="blob-num js-line-number" data-line-number="187"></td>
        <td id="LC187" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">elif</span> c <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span><span class="pl-cce">\\</span><span class="pl-pds">&#39;</span></span>:</td>
      </tr>
      <tr>
        <td id="L188" class="blob-num js-line-number" data-line-number="188"></td>
        <td id="LC188" class="blob-code blob-code-inner js-file-line">				<span class="pl-k">if</span> index <span class="pl-k">==</span> json.Length:</td>
      </tr>
      <tr>
        <td id="L189" class="blob-num js-line-number" data-line-number="189"></td>
        <td id="LC189" class="blob-code blob-code-inner js-file-line">					<span class="pl-k">break</span></td>
      </tr>
      <tr>
        <td id="L190" class="blob-num js-line-number" data-line-number="190"></td>
        <td id="LC190" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L191" class="blob-num js-line-number" data-line-number="191"></td>
        <td id="LC191" class="blob-code blob-code-inner js-file-line">				c <span class="pl-k">=</span> json[index]</td>
      </tr>
      <tr>
        <td id="L192" class="blob-num js-line-number" data-line-number="192"></td>
        <td id="LC192" class="blob-code blob-code-inner js-file-line">				index <span class="pl-k">+=</span> <span class="pl-c1">1</span></td>
      </tr>
      <tr>
        <td id="L193" class="blob-num js-line-number" data-line-number="193"></td>
        <td id="LC193" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L194" class="blob-num js-line-number" data-line-number="194"></td>
        <td id="LC194" class="blob-code blob-code-inner js-file-line">				<span class="pl-k">if</span> c <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>&quot;<span class="pl-pds">&#39;</span></span>:</td>
      </tr>
      <tr>
        <td id="L195" class="blob-num js-line-number" data-line-number="195"></td>
        <td id="LC195" class="blob-code blob-code-inner js-file-line">					s.Append(<span class="pl-s"><span class="pl-pds">&#39;</span>&quot;<span class="pl-pds">&#39;</span></span>)</td>
      </tr>
      <tr>
        <td id="L196" class="blob-num js-line-number" data-line-number="196"></td>
        <td id="LC196" class="blob-code blob-code-inner js-file-line">				<span class="pl-k">elif</span> c <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span><span class="pl-cce">\\</span><span class="pl-pds">&#39;</span></span>:</td>
      </tr>
      <tr>
        <td id="L197" class="blob-num js-line-number" data-line-number="197"></td>
        <td id="LC197" class="blob-code blob-code-inner js-file-line">					s.Append(<span class="pl-s"><span class="pl-pds">&#39;</span><span class="pl-cce">\\</span><span class="pl-pds">&#39;</span></span>)</td>
      </tr>
      <tr>
        <td id="L198" class="blob-num js-line-number" data-line-number="198"></td>
        <td id="LC198" class="blob-code blob-code-inner js-file-line">				<span class="pl-k">elif</span> c <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>/<span class="pl-pds">&#39;</span></span>:</td>
      </tr>
      <tr>
        <td id="L199" class="blob-num js-line-number" data-line-number="199"></td>
        <td id="LC199" class="blob-code blob-code-inner js-file-line">					s.Append(<span class="pl-s"><span class="pl-pds">&#39;</span>/<span class="pl-pds">&#39;</span></span>)</td>
      </tr>
      <tr>
        <td id="L200" class="blob-num js-line-number" data-line-number="200"></td>
        <td id="LC200" class="blob-code blob-code-inner js-file-line">				<span class="pl-k">elif</span> c <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>b<span class="pl-pds">&#39;</span></span>:</td>
      </tr>
      <tr>
        <td id="L201" class="blob-num js-line-number" data-line-number="201"></td>
        <td id="LC201" class="blob-code blob-code-inner js-file-line">					s.Append(<span class="pl-s"><span class="pl-pds">&#39;</span><span class="pl-cce">\b</span><span class="pl-pds">&#39;</span></span>)</td>
      </tr>
      <tr>
        <td id="L202" class="blob-num js-line-number" data-line-number="202"></td>
        <td id="LC202" class="blob-code blob-code-inner js-file-line">				<span class="pl-k">elif</span> c <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>f<span class="pl-pds">&#39;</span></span>:</td>
      </tr>
      <tr>
        <td id="L203" class="blob-num js-line-number" data-line-number="203"></td>
        <td id="LC203" class="blob-code blob-code-inner js-file-line">					s.Append(<span class="pl-s"><span class="pl-pds">&#39;</span><span class="pl-cce">\f</span><span class="pl-pds">&#39;</span></span>)</td>
      </tr>
      <tr>
        <td id="L204" class="blob-num js-line-number" data-line-number="204"></td>
        <td id="LC204" class="blob-code blob-code-inner js-file-line">				<span class="pl-k">elif</span> c <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>n<span class="pl-pds">&#39;</span></span>:</td>
      </tr>
      <tr>
        <td id="L205" class="blob-num js-line-number" data-line-number="205"></td>
        <td id="LC205" class="blob-code blob-code-inner js-file-line">					s.Append(<span class="pl-s"><span class="pl-pds">&#39;</span><span class="pl-cce">\n</span><span class="pl-pds">&#39;</span></span>)</td>
      </tr>
      <tr>
        <td id="L206" class="blob-num js-line-number" data-line-number="206"></td>
        <td id="LC206" class="blob-code blob-code-inner js-file-line">				<span class="pl-k">elif</span> c <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>r<span class="pl-pds">&#39;</span></span>:</td>
      </tr>
      <tr>
        <td id="L207" class="blob-num js-line-number" data-line-number="207"></td>
        <td id="LC207" class="blob-code blob-code-inner js-file-line">					s.Append(<span class="pl-s"><span class="pl-pds">&#39;</span><span class="pl-cce">\r</span><span class="pl-pds">&#39;</span></span>)</td>
      </tr>
      <tr>
        <td id="L208" class="blob-num js-line-number" data-line-number="208"></td>
        <td id="LC208" class="blob-code blob-code-inner js-file-line">				<span class="pl-k">elif</span> c <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>t<span class="pl-pds">&#39;</span></span>:</td>
      </tr>
      <tr>
        <td id="L209" class="blob-num js-line-number" data-line-number="209"></td>
        <td id="LC209" class="blob-code blob-code-inner js-file-line">					s.Append(<span class="pl-s"><span class="pl-pds">&#39;</span><span class="pl-cce">\t</span><span class="pl-pds">&#39;</span></span>)</td>
      </tr>
      <tr>
        <td id="L210" class="blob-num js-line-number" data-line-number="210"></td>
        <td id="LC210" class="blob-code blob-code-inner js-file-line">				<span class="pl-k">elif</span> c <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>u<span class="pl-pds">&#39;</span></span>:</td>
      </tr>
      <tr>
        <td id="L211" class="blob-num js-line-number" data-line-number="211"></td>
        <td id="LC211" class="blob-code blob-code-inner js-file-line">					remainingLength <span class="pl-k">=</span> json.Length <span class="pl-k">-</span> index</td>
      </tr>
      <tr>
        <td id="L212" class="blob-num js-line-number" data-line-number="212"></td>
        <td id="LC212" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L213" class="blob-num js-line-number" data-line-number="213"></td>
        <td id="LC213" class="blob-code blob-code-inner js-file-line">					<span class="pl-k">if</span> remainingLength <span class="pl-k">&gt;=</span> <span class="pl-c1">4</span>:</td>
      </tr>
      <tr>
        <td id="L214" class="blob-num js-line-number" data-line-number="214"></td>
        <td id="LC214" class="blob-code blob-code-inner js-file-line">						sb <span class="pl-k">=</span> StringBuilder()</td>
      </tr>
      <tr>
        <td id="L215" class="blob-num js-line-number" data-line-number="215"></td>
        <td id="LC215" class="blob-code blob-code-inner js-file-line">						</td>
      </tr>
      <tr>
        <td id="L216" class="blob-num js-line-number" data-line-number="216"></td>
        <td id="LC216" class="blob-code blob-code-inner js-file-line">						<span class="pl-k">for</span> i <span class="pl-k">in</span> <span class="pl-c1">range</span>(<span class="pl-c1">4</span>):</td>
      </tr>
      <tr>
        <td id="L217" class="blob-num js-line-number" data-line-number="217"></td>
        <td id="LC217" class="blob-code blob-code-inner js-file-line">							sb.Append(json[index <span class="pl-k">+</span> i])</td>
      </tr>
      <tr>
        <td id="L218" class="blob-num js-line-number" data-line-number="218"></td>
        <td id="LC218" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L219" class="blob-num js-line-number" data-line-number="219"></td>
        <td id="LC219" class="blob-code blob-code-inner js-file-line">						success, codePoint <span class="pl-k">=</span> UInt32.TryParse(sb.ToString(), NumberStyles.HexNumber, CultureInfo.InvariantCulture)</td>
      </tr>
      <tr>
        <td id="L220" class="blob-num js-line-number" data-line-number="220"></td>
        <td id="LC220" class="blob-code blob-code-inner js-file-line">						</td>
      </tr>
      <tr>
        <td id="L221" class="blob-num js-line-number" data-line-number="221"></td>
        <td id="LC221" class="blob-code blob-code-inner js-file-line">						<span class="pl-k">if</span> <span class="pl-k">not</span> success:</td>
      </tr>
      <tr>
        <td id="L222" class="blob-num js-line-number" data-line-number="222"></td>
        <td id="LC222" class="blob-code blob-code-inner js-file-line">							<span class="pl-k">return</span> String.Empty, index, success</td>
      </tr>
      <tr>
        <td id="L223" class="blob-num js-line-number" data-line-number="223"></td>
        <td id="LC223" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L224" class="blob-num js-line-number" data-line-number="224"></td>
        <td id="LC224" class="blob-code blob-code-inner js-file-line">						s.Append(Encoding.UTF32.GetString(BitConverter.GetBytes(codePoint)))</td>
      </tr>
      <tr>
        <td id="L225" class="blob-num js-line-number" data-line-number="225"></td>
        <td id="LC225" class="blob-code blob-code-inner js-file-line">						index <span class="pl-k">+=</span> <span class="pl-c1">4</span></td>
      </tr>
      <tr>
        <td id="L226" class="blob-num js-line-number" data-line-number="226"></td>
        <td id="LC226" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L227" class="blob-num js-line-number" data-line-number="227"></td>
        <td id="LC227" class="blob-code blob-code-inner js-file-line">					<span class="pl-k">else</span>:</td>
      </tr>
      <tr>
        <td id="L228" class="blob-num js-line-number" data-line-number="228"></td>
        <td id="LC228" class="blob-code blob-code-inner js-file-line">						<span class="pl-k">break</span></td>
      </tr>
      <tr>
        <td id="L229" class="blob-num js-line-number" data-line-number="229"></td>
        <td id="LC229" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L230" class="blob-num js-line-number" data-line-number="230"></td>
        <td id="LC230" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">else</span>:</td>
      </tr>
      <tr>
        <td id="L231" class="blob-num js-line-number" data-line-number="231"></td>
        <td id="LC231" class="blob-code blob-code-inner js-file-line">				s.Append(c)</td>
      </tr>
      <tr>
        <td id="L232" class="blob-num js-line-number" data-line-number="232"></td>
        <td id="LC232" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L233" class="blob-num js-line-number" data-line-number="233"></td>
        <td id="LC233" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">if</span> <span class="pl-k">not</span> complete:</td>
      </tr>
      <tr>
        <td id="L234" class="blob-num js-line-number" data-line-number="234"></td>
        <td id="LC234" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">return</span> <span class="pl-c1">None</span>, index, <span class="pl-c1">False</span></td>
      </tr>
      <tr>
        <td id="L235" class="blob-num js-line-number" data-line-number="235"></td>
        <td id="LC235" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L236" class="blob-num js-line-number" data-line-number="236"></td>
        <td id="LC236" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">return</span> s.ToString(), index, success</td>
      </tr>
      <tr>
        <td id="L237" class="blob-num js-line-number" data-line-number="237"></td>
        <td id="LC237" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L238" class="blob-num js-line-number" data-line-number="238"></td>
        <td id="LC238" class="blob-code blob-code-inner js-file-line">	<span class="pl-en">@<span class="pl-c1">staticmethod</span></span></td>
      </tr>
      <tr>
        <td id="L239" class="blob-num js-line-number" data-line-number="239"></td>
        <td id="LC239" class="blob-code blob-code-inner js-file-line">	<span class="pl-k">def</span> <span class="pl-en">parseNumber</span>(<span class="pl-smi">json</span>, <span class="pl-smi">index</span>, <span class="pl-smi">success</span>):</td>
      </tr>
      <tr>
        <td id="L240" class="blob-num js-line-number" data-line-number="240"></td>
        <td id="LC240" class="blob-code blob-code-inner js-file-line">		index <span class="pl-k">=</span> JsonDecoder.skipWhitespace(json, index)</td>
      </tr>
      <tr>
        <td id="L241" class="blob-num js-line-number" data-line-number="241"></td>
        <td id="LC241" class="blob-code blob-code-inner js-file-line">		lastIndex <span class="pl-k">=</span> JsonDecoder.getLastIndexOfNumber(json, index)</td>
      </tr>
      <tr>
        <td id="L242" class="blob-num js-line-number" data-line-number="242"></td>
        <td id="LC242" class="blob-code blob-code-inner js-file-line">		charLength <span class="pl-k">=</span> (lastIndex <span class="pl-k">-</span> index) <span class="pl-k">+</span> <span class="pl-c1">1</span></td>
      </tr>
      <tr>
        <td id="L243" class="blob-num js-line-number" data-line-number="243"></td>
        <td id="LC243" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L244" class="blob-num js-line-number" data-line-number="244"></td>
        <td id="LC244" class="blob-code blob-code-inner js-file-line">		sb <span class="pl-k">=</span> StringBuilder()</td>
      </tr>
      <tr>
        <td id="L245" class="blob-num js-line-number" data-line-number="245"></td>
        <td id="LC245" class="blob-code blob-code-inner js-file-line">						</td>
      </tr>
      <tr>
        <td id="L246" class="blob-num js-line-number" data-line-number="246"></td>
        <td id="LC246" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">for</span> i <span class="pl-k">in</span> <span class="pl-c1">range</span>(charLength):</td>
      </tr>
      <tr>
        <td id="L247" class="blob-num js-line-number" data-line-number="247"></td>
        <td id="LC247" class="blob-code blob-code-inner js-file-line">			sb.Append(json[index <span class="pl-k">+</span> i])</td>
      </tr>
      <tr>
        <td id="L248" class="blob-num js-line-number" data-line-number="248"></td>
        <td id="LC248" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L249" class="blob-num js-line-number" data-line-number="249"></td>
        <td id="LC249" class="blob-code blob-code-inner js-file-line">		success, number <span class="pl-k">=</span> Double.TryParse(sb.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture)</td>
      </tr>
      <tr>
        <td id="L250" class="blob-num js-line-number" data-line-number="250"></td>
        <td id="LC250" class="blob-code blob-code-inner js-file-line">		index <span class="pl-k">=</span> lastIndex <span class="pl-k">+</span> <span class="pl-c1">1</span></td>
      </tr>
      <tr>
        <td id="L251" class="blob-num js-line-number" data-line-number="251"></td>
        <td id="LC251" class="blob-code blob-code-inner js-file-line">		</td>
      </tr>
      <tr>
        <td id="L252" class="blob-num js-line-number" data-line-number="252"></td>
        <td id="LC252" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">return</span> number, index, success</td>
      </tr>
      <tr>
        <td id="L253" class="blob-num js-line-number" data-line-number="253"></td>
        <td id="LC253" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L254" class="blob-num js-line-number" data-line-number="254"></td>
        <td id="LC254" class="blob-code blob-code-inner js-file-line">	<span class="pl-en">@<span class="pl-c1">staticmethod</span></span></td>
      </tr>
      <tr>
        <td id="L255" class="blob-num js-line-number" data-line-number="255"></td>
        <td id="LC255" class="blob-code blob-code-inner js-file-line">	<span class="pl-k">def</span> <span class="pl-en">getLastIndexOfNumber</span>(<span class="pl-smi">json</span>, <span class="pl-smi">index</span>):</td>
      </tr>
      <tr>
        <td id="L256" class="blob-num js-line-number" data-line-number="256"></td>
        <td id="LC256" class="blob-code blob-code-inner js-file-line">		lastIndex <span class="pl-k">=</span> index</td>
      </tr>
      <tr>
        <td id="L257" class="blob-num js-line-number" data-line-number="257"></td>
        <td id="LC257" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L258" class="blob-num js-line-number" data-line-number="258"></td>
        <td id="LC258" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">while</span> lastIndex <span class="pl-k">&lt;</span> json.Length:</td>
      </tr>
      <tr>
        <td id="L259" class="blob-num js-line-number" data-line-number="259"></td>
        <td id="LC259" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">if</span> <span class="pl-s"><span class="pl-pds">&quot;</span>0123456789+-.eE<span class="pl-pds">&quot;</span></span>.IndexOf(json[lastIndex]) <span class="pl-k">==</span> <span class="pl-k">-</span><span class="pl-c1">1</span>:</td>
      </tr>
      <tr>
        <td id="L260" class="blob-num js-line-number" data-line-number="260"></td>
        <td id="LC260" class="blob-code blob-code-inner js-file-line">				<span class="pl-k">break</span></td>
      </tr>
      <tr>
        <td id="L261" class="blob-num js-line-number" data-line-number="261"></td>
        <td id="LC261" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L262" class="blob-num js-line-number" data-line-number="262"></td>
        <td id="LC262" class="blob-code blob-code-inner js-file-line">			lastIndex <span class="pl-k">+=</span> <span class="pl-c1">1</span></td>
      </tr>
      <tr>
        <td id="L263" class="blob-num js-line-number" data-line-number="263"></td>
        <td id="LC263" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L264" class="blob-num js-line-number" data-line-number="264"></td>
        <td id="LC264" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">return</span> lastIndex <span class="pl-k">-</span> <span class="pl-c1">1</span></td>
      </tr>
      <tr>
        <td id="L265" class="blob-num js-line-number" data-line-number="265"></td>
        <td id="LC265" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L266" class="blob-num js-line-number" data-line-number="266"></td>
        <td id="LC266" class="blob-code blob-code-inner js-file-line">	<span class="pl-en">@<span class="pl-c1">staticmethod</span></span></td>
      </tr>
      <tr>
        <td id="L267" class="blob-num js-line-number" data-line-number="267"></td>
        <td id="LC267" class="blob-code blob-code-inner js-file-line">	<span class="pl-k">def</span> <span class="pl-en">skipWhitespace</span>(<span class="pl-smi">json</span>, <span class="pl-smi">index</span>):</td>
      </tr>
      <tr>
        <td id="L268" class="blob-num js-line-number" data-line-number="268"></td>
        <td id="LC268" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">while</span> index <span class="pl-k">&lt;</span> json.Length:</td>
      </tr>
      <tr>
        <td id="L269" class="blob-num js-line-number" data-line-number="269"></td>
        <td id="LC269" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">if</span> <span class="pl-s"><span class="pl-pds">&quot;</span> <span class="pl-cce">\t</span><span class="pl-cce">\n</span><span class="pl-cce">\r</span><span class="pl-pds">&quot;</span></span>.IndexOf(json[index]) <span class="pl-k">==</span> <span class="pl-k">-</span><span class="pl-c1">1</span>:</td>
      </tr>
      <tr>
        <td id="L270" class="blob-num js-line-number" data-line-number="270"></td>
        <td id="LC270" class="blob-code blob-code-inner js-file-line">				<span class="pl-k">break</span></td>
      </tr>
      <tr>
        <td id="L271" class="blob-num js-line-number" data-line-number="271"></td>
        <td id="LC271" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L272" class="blob-num js-line-number" data-line-number="272"></td>
        <td id="LC272" class="blob-code blob-code-inner js-file-line">			index <span class="pl-k">+=</span> <span class="pl-c1">1</span></td>
      </tr>
      <tr>
        <td id="L273" class="blob-num js-line-number" data-line-number="273"></td>
        <td id="LC273" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L274" class="blob-num js-line-number" data-line-number="274"></td>
        <td id="LC274" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">return</span> index</td>
      </tr>
      <tr>
        <td id="L275" class="blob-num js-line-number" data-line-number="275"></td>
        <td id="LC275" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L276" class="blob-num js-line-number" data-line-number="276"></td>
        <td id="LC276" class="blob-code blob-code-inner js-file-line">	<span class="pl-en">@<span class="pl-c1">staticmethod</span></span></td>
      </tr>
      <tr>
        <td id="L277" class="blob-num js-line-number" data-line-number="277"></td>
        <td id="LC277" class="blob-code blob-code-inner js-file-line">	<span class="pl-k">def</span> <span class="pl-en">lookAhead</span>(<span class="pl-smi">json</span>, <span class="pl-smi">index</span>):</td>
      </tr>
      <tr>
        <td id="L278" class="blob-num js-line-number" data-line-number="278"></td>
        <td id="LC278" class="blob-code blob-code-inner js-file-line">		i, j <span class="pl-k">=</span> JsonDecoder.nextToken(json, index)</td>
      </tr>
      <tr>
        <td id="L279" class="blob-num js-line-number" data-line-number="279"></td>
        <td id="LC279" class="blob-code blob-code-inner js-file-line">		</td>
      </tr>
      <tr>
        <td id="L280" class="blob-num js-line-number" data-line-number="280"></td>
        <td id="LC280" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">return</span> i</td>
      </tr>
      <tr>
        <td id="L281" class="blob-num js-line-number" data-line-number="281"></td>
        <td id="LC281" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L282" class="blob-num js-line-number" data-line-number="282"></td>
        <td id="LC282" class="blob-code blob-code-inner js-file-line">	<span class="pl-en">@<span class="pl-c1">staticmethod</span></span></td>
      </tr>
      <tr>
        <td id="L283" class="blob-num js-line-number" data-line-number="283"></td>
        <td id="LC283" class="blob-code blob-code-inner js-file-line">	<span class="pl-k">def</span> <span class="pl-en">nextToken</span>(<span class="pl-smi">json</span>, <span class="pl-smi">index</span>):</td>
      </tr>
      <tr>
        <td id="L284" class="blob-num js-line-number" data-line-number="284"></td>
        <td id="LC284" class="blob-code blob-code-inner js-file-line">		index <span class="pl-k">=</span> JsonDecoder.skipWhitespace(json, index)</td>
      </tr>
      <tr>
        <td id="L285" class="blob-num js-line-number" data-line-number="285"></td>
        <td id="LC285" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L286" class="blob-num js-line-number" data-line-number="286"></td>
        <td id="LC286" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">if</span> index <span class="pl-k">==</span> json.Length:</td>
      </tr>
      <tr>
        <td id="L287" class="blob-num js-line-number" data-line-number="287"></td>
        <td id="LC287" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">return</span> JsonDecoder.TOKEN_NONE, index</td>
      </tr>
      <tr>
        <td id="L288" class="blob-num js-line-number" data-line-number="288"></td>
        <td id="LC288" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L289" class="blob-num js-line-number" data-line-number="289"></td>
        <td id="LC289" class="blob-code blob-code-inner js-file-line">		c <span class="pl-k">=</span> json[index]</td>
      </tr>
      <tr>
        <td id="L290" class="blob-num js-line-number" data-line-number="290"></td>
        <td id="LC290" class="blob-code blob-code-inner js-file-line">		index <span class="pl-k">+=</span> <span class="pl-c1">1</span></td>
      </tr>
      <tr>
        <td id="L291" class="blob-num js-line-number" data-line-number="291"></td>
        <td id="LC291" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L292" class="blob-num js-line-number" data-line-number="292"></td>
        <td id="LC292" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">if</span> c <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>{<span class="pl-pds">&#39;</span></span>:</td>
      </tr>
      <tr>
        <td id="L293" class="blob-num js-line-number" data-line-number="293"></td>
        <td id="LC293" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">return</span> JsonDecoder.TOKEN_CURLY_OPEN, index</td>
      </tr>
      <tr>
        <td id="L294" class="blob-num js-line-number" data-line-number="294"></td>
        <td id="LC294" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L295" class="blob-num js-line-number" data-line-number="295"></td>
        <td id="LC295" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">elif</span> c <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>}<span class="pl-pds">&#39;</span></span>:</td>
      </tr>
      <tr>
        <td id="L296" class="blob-num js-line-number" data-line-number="296"></td>
        <td id="LC296" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">return</span> JsonDecoder.TOKEN_CURLY_CLOSE, index</td>
      </tr>
      <tr>
        <td id="L297" class="blob-num js-line-number" data-line-number="297"></td>
        <td id="LC297" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L298" class="blob-num js-line-number" data-line-number="298"></td>
        <td id="LC298" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">elif</span> c <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>[<span class="pl-pds">&#39;</span></span>:</td>
      </tr>
      <tr>
        <td id="L299" class="blob-num js-line-number" data-line-number="299"></td>
        <td id="LC299" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">return</span> JsonDecoder.TOKEN_SQUARED_OPEN, index</td>
      </tr>
      <tr>
        <td id="L300" class="blob-num js-line-number" data-line-number="300"></td>
        <td id="LC300" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L301" class="blob-num js-line-number" data-line-number="301"></td>
        <td id="LC301" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">elif</span> c <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>]<span class="pl-pds">&#39;</span></span>:</td>
      </tr>
      <tr>
        <td id="L302" class="blob-num js-line-number" data-line-number="302"></td>
        <td id="LC302" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">return</span> JsonDecoder.TOKEN_SQUARED_CLOSE, index</td>
      </tr>
      <tr>
        <td id="L303" class="blob-num js-line-number" data-line-number="303"></td>
        <td id="LC303" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L304" class="blob-num js-line-number" data-line-number="304"></td>
        <td id="LC304" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">elif</span> c <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>,<span class="pl-pds">&#39;</span></span>:</td>
      </tr>
      <tr>
        <td id="L305" class="blob-num js-line-number" data-line-number="305"></td>
        <td id="LC305" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">return</span> JsonDecoder.TOKEN_COMMA, index</td>
      </tr>
      <tr>
        <td id="L306" class="blob-num js-line-number" data-line-number="306"></td>
        <td id="LC306" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L307" class="blob-num js-line-number" data-line-number="307"></td>
        <td id="LC307" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">elif</span> c <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>&quot;<span class="pl-pds">&#39;</span></span>:</td>
      </tr>
      <tr>
        <td id="L308" class="blob-num js-line-number" data-line-number="308"></td>
        <td id="LC308" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">return</span> JsonDecoder.TOKEN_STRING, index</td>
      </tr>
      <tr>
        <td id="L309" class="blob-num js-line-number" data-line-number="309"></td>
        <td id="LC309" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L310" class="blob-num js-line-number" data-line-number="310"></td>
        <td id="LC310" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">elif</span> c <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>0<span class="pl-pds">&#39;</span></span> <span class="pl-k">or</span> c <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>1<span class="pl-pds">&#39;</span></span> <span class="pl-k">or</span> c <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>2<span class="pl-pds">&#39;</span></span> <span class="pl-k">or</span> c <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>3<span class="pl-pds">&#39;</span></span> <span class="pl-k">or</span> c <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>4<span class="pl-pds">&#39;</span></span> <span class="pl-k">or</span> c <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>5<span class="pl-pds">&#39;</span></span> <span class="pl-k">or</span> c <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>6<span class="pl-pds">&#39;</span></span> <span class="pl-k">or</span> c <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>7<span class="pl-pds">&#39;</span></span> <span class="pl-k">or</span> c <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>8<span class="pl-pds">&#39;</span></span> <span class="pl-k">or</span> c <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>9<span class="pl-pds">&#39;</span></span> <span class="pl-k">or</span> c <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>-<span class="pl-pds">&#39;</span></span>:</td>
      </tr>
      <tr>
        <td id="L311" class="blob-num js-line-number" data-line-number="311"></td>
        <td id="LC311" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">return</span> JsonDecoder.TOKEN_NUMBER, index</td>
      </tr>
      <tr>
        <td id="L312" class="blob-num js-line-number" data-line-number="312"></td>
        <td id="LC312" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L313" class="blob-num js-line-number" data-line-number="313"></td>
        <td id="LC313" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">elif</span> c <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>:<span class="pl-pds">&#39;</span></span>:</td>
      </tr>
      <tr>
        <td id="L314" class="blob-num js-line-number" data-line-number="314"></td>
        <td id="LC314" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">return</span> JsonDecoder.TOKEN_COLON, index</td>
      </tr>
      <tr>
        <td id="L315" class="blob-num js-line-number" data-line-number="315"></td>
        <td id="LC315" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L316" class="blob-num js-line-number" data-line-number="316"></td>
        <td id="LC316" class="blob-code blob-code-inner js-file-line">		index <span class="pl-k">-=</span> <span class="pl-c1">1</span></td>
      </tr>
      <tr>
        <td id="L317" class="blob-num js-line-number" data-line-number="317"></td>
        <td id="LC317" class="blob-code blob-code-inner js-file-line">		remainingLength <span class="pl-k">=</span> json.Length <span class="pl-k">-</span> index</td>
      </tr>
      <tr>
        <td id="L318" class="blob-num js-line-number" data-line-number="318"></td>
        <td id="LC318" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L319" class="blob-num js-line-number" data-line-number="319"></td>
        <td id="LC319" class="blob-code blob-code-inner js-file-line">		<span class="pl-c"># false</span></td>
      </tr>
      <tr>
        <td id="L320" class="blob-num js-line-number" data-line-number="320"></td>
        <td id="LC320" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">if</span> remainingLength <span class="pl-k">&gt;=</span> <span class="pl-c1">5</span>:</td>
      </tr>
      <tr>
        <td id="L321" class="blob-num js-line-number" data-line-number="321"></td>
        <td id="LC321" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">if</span> json[index] <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>f<span class="pl-pds">&#39;</span></span> <span class="pl-k">and</span> json[index <span class="pl-k">+</span> <span class="pl-c1">1</span>] <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>a<span class="pl-pds">&#39;</span></span> <span class="pl-k">and</span> json[index <span class="pl-k">+</span> <span class="pl-c1">2</span>] <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>l<span class="pl-pds">&#39;</span></span> <span class="pl-k">and</span> json[index <span class="pl-k">+</span> <span class="pl-c1">3</span>] <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>s<span class="pl-pds">&#39;</span></span> <span class="pl-k">and</span> json[index <span class="pl-k">+</span> <span class="pl-c1">4</span>] <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>e<span class="pl-pds">&#39;</span></span>:</td>
      </tr>
      <tr>
        <td id="L322" class="blob-num js-line-number" data-line-number="322"></td>
        <td id="LC322" class="blob-code blob-code-inner js-file-line">				index <span class="pl-k">+=</span> <span class="pl-c1">5</span></td>
      </tr>
      <tr>
        <td id="L323" class="blob-num js-line-number" data-line-number="323"></td>
        <td id="LC323" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L324" class="blob-num js-line-number" data-line-number="324"></td>
        <td id="LC324" class="blob-code blob-code-inner js-file-line">				<span class="pl-k">return</span> JsonDecoder.TOKEN_FALSE, index</td>
      </tr>
      <tr>
        <td id="L325" class="blob-num js-line-number" data-line-number="325"></td>
        <td id="LC325" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L326" class="blob-num js-line-number" data-line-number="326"></td>
        <td id="LC326" class="blob-code blob-code-inner js-file-line">		<span class="pl-c"># true</span></td>
      </tr>
      <tr>
        <td id="L327" class="blob-num js-line-number" data-line-number="327"></td>
        <td id="LC327" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">if</span> remainingLength <span class="pl-k">&gt;=</span> <span class="pl-c1">4</span>:</td>
      </tr>
      <tr>
        <td id="L328" class="blob-num js-line-number" data-line-number="328"></td>
        <td id="LC328" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">if</span> json[index] <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>t<span class="pl-pds">&#39;</span></span> <span class="pl-k">and</span> json[index <span class="pl-k">+</span> <span class="pl-c1">1</span>] <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>r<span class="pl-pds">&#39;</span></span> <span class="pl-k">and</span> json[index <span class="pl-k">+</span> <span class="pl-c1">2</span>] <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>u<span class="pl-pds">&#39;</span></span> <span class="pl-k">and</span> json[index <span class="pl-k">+</span> <span class="pl-c1">3</span>] <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>e<span class="pl-pds">&#39;</span></span>:</td>
      </tr>
      <tr>
        <td id="L329" class="blob-num js-line-number" data-line-number="329"></td>
        <td id="LC329" class="blob-code blob-code-inner js-file-line">				index <span class="pl-k">+=</span> <span class="pl-c1">4</span></td>
      </tr>
      <tr>
        <td id="L330" class="blob-num js-line-number" data-line-number="330"></td>
        <td id="LC330" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L331" class="blob-num js-line-number" data-line-number="331"></td>
        <td id="LC331" class="blob-code blob-code-inner js-file-line">				<span class="pl-k">return</span> JsonDecoder.TOKEN_TRUE, index</td>
      </tr>
      <tr>
        <td id="L332" class="blob-num js-line-number" data-line-number="332"></td>
        <td id="LC332" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L333" class="blob-num js-line-number" data-line-number="333"></td>
        <td id="LC333" class="blob-code blob-code-inner js-file-line">		<span class="pl-c"># null</span></td>
      </tr>
      <tr>
        <td id="L334" class="blob-num js-line-number" data-line-number="334"></td>
        <td id="LC334" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">if</span> remainingLength <span class="pl-k">&gt;=</span> <span class="pl-c1">4</span>:</td>
      </tr>
      <tr>
        <td id="L335" class="blob-num js-line-number" data-line-number="335"></td>
        <td id="LC335" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">if</span> json[index] <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>n<span class="pl-pds">&#39;</span></span> <span class="pl-k">and</span> json[index <span class="pl-k">+</span> <span class="pl-c1">1</span>] <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>u<span class="pl-pds">&#39;</span></span> <span class="pl-k">and</span> json[index <span class="pl-k">+</span> <span class="pl-c1">2</span>] <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>l<span class="pl-pds">&#39;</span></span> <span class="pl-k">and</span> json[index <span class="pl-k">+</span> <span class="pl-c1">3</span>] <span class="pl-k">==</span> <span class="pl-s"><span class="pl-pds">&#39;</span>l<span class="pl-pds">&#39;</span></span>:</td>
      </tr>
      <tr>
        <td id="L336" class="blob-num js-line-number" data-line-number="336"></td>
        <td id="LC336" class="blob-code blob-code-inner js-file-line">				index <span class="pl-k">+=</span> <span class="pl-c1">4</span></td>
      </tr>
      <tr>
        <td id="L337" class="blob-num js-line-number" data-line-number="337"></td>
        <td id="LC337" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L338" class="blob-num js-line-number" data-line-number="338"></td>
        <td id="LC338" class="blob-code blob-code-inner js-file-line">				<span class="pl-k">return</span> JsonDecoder.TOKEN_NULL, index</td>
      </tr>
      <tr>
        <td id="L339" class="blob-num js-line-number" data-line-number="339"></td>
        <td id="LC339" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L340" class="blob-num js-line-number" data-line-number="340"></td>
        <td id="LC340" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">return</span> JsonDecoder.TOKEN_NONE, index</td>
      </tr>
      <tr>
        <td id="L341" class="blob-num js-line-number" data-line-number="341"></td>
        <td id="LC341" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L342" class="blob-num js-line-number" data-line-number="342"></td>
        <td id="LC342" class="blob-code blob-code-inner js-file-line"><span class="pl-k">def</span> <span class="pl-en">onTick</span>(<span class="pl-smi">timer</span>, <span class="pl-smi">e</span>):</td>
      </tr>
      <tr>
        <td id="L343" class="blob-num js-line-number" data-line-number="343"></td>
        <td id="LC343" class="blob-code blob-code-inner js-file-line">	<span class="pl-k">def</span> <span class="pl-en">onUpdate</span>():</td>
      </tr>
      <tr>
        <td id="L344" class="blob-num js-line-number" data-line-number="344"></td>
        <td id="LC344" class="blob-code blob-code-inner js-file-line">		temp <span class="pl-k">=</span> <span class="pl-c1">0</span></td>
      </tr>
      <tr>
        <td id="L345" class="blob-num js-line-number" data-line-number="345"></td>
        <td id="LC345" class="blob-code blob-code-inner js-file-line">		windSpeed <span class="pl-k">=</span> <span class="pl-c1">0</span></td>
      </tr>
      <tr>
        <td id="L346" class="blob-num js-line-number" data-line-number="346"></td>
        <td id="LC346" class="blob-code blob-code-inner js-file-line">		windDeg <span class="pl-k">=</span> <span class="pl-c1">0</span></td>
      </tr>
      <tr>
        <td id="L347" class="blob-num js-line-number" data-line-number="347"></td>
        <td id="LC347" class="blob-code blob-code-inner js-file-line">		weatherIdList <span class="pl-k">=</span> List[Double]()</td>
      </tr>
      <tr>
        <td id="L348" class="blob-num js-line-number" data-line-number="348"></td>
        <td id="LC348" class="blob-code blob-code-inner js-file-line">		weatherPathHashSet <span class="pl-k">=</span> HashSet[String]()</td>
      </tr>
      <tr>
        <td id="L349" class="blob-num js-line-number" data-line-number="349"></td>
        <td id="LC349" class="blob-code blob-code-inner js-file-line">		weatherStreamList <span class="pl-k">=</span> List[MemoryStream]()</td>
      </tr>
      <tr>
        <td id="L350" class="blob-num js-line-number" data-line-number="350"></td>
        <td id="LC350" class="blob-code blob-code-inner js-file-line">		weatherConditionList <span class="pl-k">=</span> List[String]()</td>
      </tr>
      <tr>
        <td id="L351" class="blob-num js-line-number" data-line-number="351"></td>
        <td id="LC351" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L352" class="blob-num js-line-number" data-line-number="352"></td>
        <td id="LC352" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">if</span> NetworkInterface.GetIsNetworkAvailable():</td>
      </tr>
      <tr>
        <td id="L353" class="blob-num js-line-number" data-line-number="353"></td>
        <td id="LC353" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">try</span>:</td>
      </tr>
      <tr>
        <td id="L354" class="blob-num js-line-number" data-line-number="354"></td>
        <td id="LC354" class="blob-code blob-code-inner js-file-line">				request <span class="pl-k">=</span> WebRequest.Create(Uri(String.Concat(<span class="pl-s"><span class="pl-pds">&quot;</span>http://api.openweathermap.org/data/2.5/find?q=<span class="pl-pds">&quot;</span></span>, urlEncode(location), <span class="pl-s"><span class="pl-pds">&quot;</span>&amp;units=metric&amp;cnt=1<span class="pl-pds">&quot;</span></span>)))</td>
      </tr>
      <tr>
        <td id="L355" class="blob-num js-line-number" data-line-number="355"></td>
        <td id="LC355" class="blob-code blob-code-inner js-file-line">				response <span class="pl-k">=</span> <span class="pl-c1">None</span></td>
      </tr>
      <tr>
        <td id="L356" class="blob-num js-line-number" data-line-number="356"></td>
        <td id="LC356" class="blob-code blob-code-inner js-file-line">				stream <span class="pl-k">=</span> <span class="pl-c1">None</span></td>
      </tr>
      <tr>
        <td id="L357" class="blob-num js-line-number" data-line-number="357"></td>
        <td id="LC357" class="blob-code blob-code-inner js-file-line">				streamReader <span class="pl-k">=</span> <span class="pl-c1">None</span></td>
      </tr>
      <tr>
        <td id="L358" class="blob-num js-line-number" data-line-number="358"></td>
        <td id="LC358" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L359" class="blob-num js-line-number" data-line-number="359"></td>
        <td id="LC359" class="blob-code blob-code-inner js-file-line">				<span class="pl-k">try</span>:</td>
      </tr>
      <tr>
        <td id="L360" class="blob-num js-line-number" data-line-number="360"></td>
        <td id="LC360" class="blob-code blob-code-inner js-file-line">					nowDateTime <span class="pl-k">=</span> DateTime.Now</td>
      </tr>
      <tr>
        <td id="L361" class="blob-num js-line-number" data-line-number="361"></td>
        <td id="LC361" class="blob-code blob-code-inner js-file-line">					response <span class="pl-k">=</span> request.GetResponse()</td>
      </tr>
      <tr>
        <td id="L362" class="blob-num js-line-number" data-line-number="362"></td>
        <td id="LC362" class="blob-code blob-code-inner js-file-line">					stream <span class="pl-k">=</span> response.GetResponseStream()</td>
      </tr>
      <tr>
        <td id="L363" class="blob-num js-line-number" data-line-number="363"></td>
        <td id="LC363" class="blob-code blob-code-inner js-file-line">					streamReader <span class="pl-k">=</span> StreamReader(stream)</td>
      </tr>
      <tr>
        <td id="L364" class="blob-num js-line-number" data-line-number="364"></td>
        <td id="LC364" class="blob-code blob-code-inner js-file-line">					jsonDictionary <span class="pl-k">=</span> JsonDecoder.decode(streamReader.ReadToEnd())</td>
      </tr>
      <tr>
        <td id="L365" class="blob-num js-line-number" data-line-number="365"></td>
        <td id="LC365" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L366" class="blob-num js-line-number" data-line-number="366"></td>
        <td id="LC366" class="blob-code blob-code-inner js-file-line">					<span class="pl-k">if</span> jsonDictionary <span class="pl-k">is</span> <span class="pl-k">not</span> <span class="pl-c1">None</span> <span class="pl-k">and</span> clr.GetClrType(Dictionary[String, Object]).IsInstanceOfType(jsonDictionary) <span class="pl-k">and</span> jsonDictionary.ContainsKey(<span class="pl-s"><span class="pl-pds">&quot;</span>list<span class="pl-pds">&quot;</span></span>) <span class="pl-k">and</span> jsonDictionary[<span class="pl-s"><span class="pl-pds">&quot;</span>list<span class="pl-pds">&quot;</span></span>] <span class="pl-k">is</span> <span class="pl-k">not</span> <span class="pl-c1">None</span> <span class="pl-k">and</span> clr.GetClrType(Array).IsInstanceOfType(jsonDictionary[<span class="pl-s"><span class="pl-pds">&quot;</span>list<span class="pl-pds">&quot;</span></span>]):</td>
      </tr>
      <tr>
        <td id="L367" class="blob-num js-line-number" data-line-number="367"></td>
        <td id="LC367" class="blob-code blob-code-inner js-file-line">						<span class="pl-k">for</span> obj <span class="pl-k">in</span> jsonDictionary[<span class="pl-s"><span class="pl-pds">&quot;</span>list<span class="pl-pds">&quot;</span></span>]:</td>
      </tr>
      <tr>
        <td id="L368" class="blob-num js-line-number" data-line-number="368"></td>
        <td id="LC368" class="blob-code blob-code-inner js-file-line">							<span class="pl-k">if</span> obj <span class="pl-k">is</span> <span class="pl-k">not</span> <span class="pl-c1">None</span> <span class="pl-k">and</span> clr.GetClrType(Dictionary[String, Object]).IsInstanceOfType(obj):</td>
      </tr>
      <tr>
        <td id="L369" class="blob-num js-line-number" data-line-number="369"></td>
        <td id="LC369" class="blob-code blob-code-inner js-file-line">								<span class="pl-k">if</span> obj.ContainsKey(<span class="pl-s"><span class="pl-pds">&quot;</span>main<span class="pl-pds">&quot;</span></span>) <span class="pl-k">and</span> obj[<span class="pl-s"><span class="pl-pds">&quot;</span>main<span class="pl-pds">&quot;</span></span>] <span class="pl-k">is</span> <span class="pl-k">not</span> <span class="pl-c1">None</span> <span class="pl-k">and</span> clr.GetClrType(Dictionary[String, Object]).IsInstanceOfType(obj[<span class="pl-s"><span class="pl-pds">&quot;</span>main<span class="pl-pds">&quot;</span></span>]) <span class="pl-k">and</span> obj[<span class="pl-s"><span class="pl-pds">&quot;</span>main<span class="pl-pds">&quot;</span></span>].ContainsKey(<span class="pl-s"><span class="pl-pds">&quot;</span>temp<span class="pl-pds">&quot;</span></span>):</td>
      </tr>
      <tr>
        <td id="L370" class="blob-num js-line-number" data-line-number="370"></td>
        <td id="LC370" class="blob-code blob-code-inner js-file-line">									temp <span class="pl-k">=</span> obj[<span class="pl-s"><span class="pl-pds">&quot;</span>main<span class="pl-pds">&quot;</span></span>][<span class="pl-s"><span class="pl-pds">&quot;</span>temp<span class="pl-pds">&quot;</span></span>]</td>
      </tr>
      <tr>
        <td id="L371" class="blob-num js-line-number" data-line-number="371"></td>
        <td id="LC371" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L372" class="blob-num js-line-number" data-line-number="372"></td>
        <td id="LC372" class="blob-code blob-code-inner js-file-line">								<span class="pl-k">if</span> obj.ContainsKey(<span class="pl-s"><span class="pl-pds">&quot;</span>wind<span class="pl-pds">&quot;</span></span>) <span class="pl-k">and</span> obj[<span class="pl-s"><span class="pl-pds">&quot;</span>wind<span class="pl-pds">&quot;</span></span>] <span class="pl-k">is</span> <span class="pl-k">not</span> <span class="pl-c1">None</span> <span class="pl-k">and</span> clr.GetClrType(Dictionary[String, Object]).IsInstanceOfType(obj[<span class="pl-s"><span class="pl-pds">&quot;</span>wind<span class="pl-pds">&quot;</span></span>]):</td>
      </tr>
      <tr>
        <td id="L373" class="blob-num js-line-number" data-line-number="373"></td>
        <td id="LC373" class="blob-code blob-code-inner js-file-line">									<span class="pl-k">if</span> obj[<span class="pl-s"><span class="pl-pds">&quot;</span>wind<span class="pl-pds">&quot;</span></span>].ContainsKey(<span class="pl-s"><span class="pl-pds">&quot;</span>speed<span class="pl-pds">&quot;</span></span>):</td>
      </tr>
      <tr>
        <td id="L374" class="blob-num js-line-number" data-line-number="374"></td>
        <td id="LC374" class="blob-code blob-code-inner js-file-line">										windSpeed <span class="pl-k">=</span> obj[<span class="pl-s"><span class="pl-pds">&quot;</span>wind<span class="pl-pds">&quot;</span></span>][<span class="pl-s"><span class="pl-pds">&quot;</span>speed<span class="pl-pds">&quot;</span></span>]</td>
      </tr>
      <tr>
        <td id="L375" class="blob-num js-line-number" data-line-number="375"></td>
        <td id="LC375" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L376" class="blob-num js-line-number" data-line-number="376"></td>
        <td id="LC376" class="blob-code blob-code-inner js-file-line">									<span class="pl-k">if</span> obj[<span class="pl-s"><span class="pl-pds">&quot;</span>wind<span class="pl-pds">&quot;</span></span>].ContainsKey(<span class="pl-s"><span class="pl-pds">&quot;</span>deg<span class="pl-pds">&quot;</span></span>):</td>
      </tr>
      <tr>
        <td id="L377" class="blob-num js-line-number" data-line-number="377"></td>
        <td id="LC377" class="blob-code blob-code-inner js-file-line">										windDeg <span class="pl-k">=</span> obj[<span class="pl-s"><span class="pl-pds">&quot;</span>wind<span class="pl-pds">&quot;</span></span>][<span class="pl-s"><span class="pl-pds">&quot;</span>deg<span class="pl-pds">&quot;</span></span>]</td>
      </tr>
      <tr>
        <td id="L378" class="blob-num js-line-number" data-line-number="378"></td>
        <td id="LC378" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L379" class="blob-num js-line-number" data-line-number="379"></td>
        <td id="LC379" class="blob-code blob-code-inner js-file-line">								<span class="pl-k">if</span> obj.ContainsKey(<span class="pl-s"><span class="pl-pds">&quot;</span>weather<span class="pl-pds">&quot;</span></span>) <span class="pl-k">and</span> obj[<span class="pl-s"><span class="pl-pds">&quot;</span>weather<span class="pl-pds">&quot;</span></span>] <span class="pl-k">is</span> <span class="pl-k">not</span> <span class="pl-c1">None</span> <span class="pl-k">and</span> clr.GetClrType(Array).IsInstanceOfType(obj[<span class="pl-s"><span class="pl-pds">&quot;</span>weather<span class="pl-pds">&quot;</span></span>]):</td>
      </tr>
      <tr>
        <td id="L380" class="blob-num js-line-number" data-line-number="380"></td>
        <td id="LC380" class="blob-code blob-code-inner js-file-line">									<span class="pl-k">for</span> o <span class="pl-k">in</span> obj[<span class="pl-s"><span class="pl-pds">&quot;</span>weather<span class="pl-pds">&quot;</span></span>]:</td>
      </tr>
      <tr>
        <td id="L381" class="blob-num js-line-number" data-line-number="381"></td>
        <td id="LC381" class="blob-code blob-code-inner js-file-line">										<span class="pl-k">if</span> o <span class="pl-k">is</span> <span class="pl-k">not</span> <span class="pl-c1">None</span> <span class="pl-k">and</span> clr.GetClrType(Dictionary[String, Object]).IsInstanceOfType(o) <span class="pl-k">and</span> o.ContainsKey(<span class="pl-s"><span class="pl-pds">&quot;</span>id<span class="pl-pds">&quot;</span></span>) <span class="pl-k">and</span> o[<span class="pl-s"><span class="pl-pds">&quot;</span>id<span class="pl-pds">&quot;</span></span>] <span class="pl-k">is</span> <span class="pl-k">not</span> <span class="pl-c1">None</span>:</td>
      </tr>
      <tr>
        <td id="L382" class="blob-num js-line-number" data-line-number="382"></td>
        <td id="LC382" class="blob-code blob-code-inner js-file-line">											weatherIdList.Add(o[<span class="pl-s"><span class="pl-pds">&quot;</span>id<span class="pl-pds">&quot;</span></span>])</td>
      </tr>
      <tr>
        <td id="L383" class="blob-num js-line-number" data-line-number="383"></td>
        <td id="LC383" class="blob-code blob-code-inner js-file-line">																</td>
      </tr>
      <tr>
        <td id="L384" class="blob-num js-line-number" data-line-number="384"></td>
        <td id="LC384" class="blob-code blob-code-inner js-file-line">								<span class="pl-k">for</span> <span class="pl-c1">id</span> <span class="pl-k">in</span> weatherIdList:</td>
      </tr>
      <tr>
        <td id="L385" class="blob-num js-line-number" data-line-number="385"></td>
        <td id="LC385" class="blob-code blob-code-inner js-file-line">									digit <span class="pl-k">=</span> Convert.ToInt32(<span class="pl-c1">id</span> <span class="pl-k">/</span> <span class="pl-c1">100</span>)</td>
      </tr>
      <tr>
        <td id="L386" class="blob-num js-line-number" data-line-number="386"></td>
        <td id="LC386" class="blob-code blob-code-inner js-file-line">									path <span class="pl-k">=</span> <span class="pl-c1">None</span></td>
      </tr>
      <tr>
        <td id="L387" class="blob-num js-line-number" data-line-number="387"></td>
        <td id="LC387" class="blob-code blob-code-inner js-file-line">									s <span class="pl-k">=</span> <span class="pl-c1">None</span></td>
      </tr>
      <tr>
        <td id="L388" class="blob-num js-line-number" data-line-number="388"></td>
        <td id="LC388" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L389" class="blob-num js-line-number" data-line-number="389"></td>
        <td id="LC389" class="blob-code blob-code-inner js-file-line">									<span class="pl-k">if</span> digit <span class="pl-k">==</span> <span class="pl-c1">2</span>:</td>
      </tr>
      <tr>
        <td id="L390" class="blob-num js-line-number" data-line-number="390"></td>
        <td id="LC390" class="blob-code blob-code-inner js-file-line">										path <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">&quot;</span>Assets<span class="pl-cce">\\</span>Cloud-Lightning.png<span class="pl-pds">&quot;</span></span></td>
      </tr>
      <tr>
        <td id="L391" class="blob-num js-line-number" data-line-number="391"></td>
        <td id="LC391" class="blob-code blob-code-inner js-file-line">										weatherConditionList.Add(<span class="pl-s"><span class="pl-pds">&quot;</span>Thunderstorm<span class="pl-pds">&quot;</span></span>)</td>
      </tr>
      <tr>
        <td id="L392" class="blob-num js-line-number" data-line-number="392"></td>
        <td id="LC392" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L393" class="blob-num js-line-number" data-line-number="393"></td>
        <td id="LC393" class="blob-code blob-code-inner js-file-line">									<span class="pl-k">elif</span>  digit <span class="pl-k">==</span> <span class="pl-c1">3</span>:</td>
      </tr>
      <tr>
        <td id="L394" class="blob-num js-line-number" data-line-number="394"></td>
        <td id="LC394" class="blob-code blob-code-inner js-file-line">										path <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">&quot;</span>Assets<span class="pl-cce">\\</span>Cloud-Drizzle.png<span class="pl-pds">&quot;</span></span></td>
      </tr>
      <tr>
        <td id="L395" class="blob-num js-line-number" data-line-number="395"></td>
        <td id="LC395" class="blob-code blob-code-inner js-file-line">										weatherConditionList.Add(<span class="pl-s"><span class="pl-pds">&quot;</span>Drizzle<span class="pl-pds">&quot;</span></span>)</td>
      </tr>
      <tr>
        <td id="L396" class="blob-num js-line-number" data-line-number="396"></td>
        <td id="LC396" class="blob-code blob-code-inner js-file-line">														</td>
      </tr>
      <tr>
        <td id="L397" class="blob-num js-line-number" data-line-number="397"></td>
        <td id="LC397" class="blob-code blob-code-inner js-file-line">									<span class="pl-k">elif</span>  digit <span class="pl-k">==</span> <span class="pl-c1">5</span>:</td>
      </tr>
      <tr>
        <td id="L398" class="blob-num js-line-number" data-line-number="398"></td>
        <td id="LC398" class="blob-code blob-code-inner js-file-line">										d <span class="pl-k">=</span> Convert.ToInt32(<span class="pl-c1">id</span> <span class="pl-k">/</span> <span class="pl-c1">10</span>)</td>
      </tr>
      <tr>
        <td id="L399" class="blob-num js-line-number" data-line-number="399"></td>
        <td id="LC399" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L400" class="blob-num js-line-number" data-line-number="400"></td>
        <td id="LC400" class="blob-code blob-code-inner js-file-line">										<span class="pl-k">if</span> d <span class="pl-k">==</span> <span class="pl-c1">0</span>:</td>
      </tr>
      <tr>
        <td id="L401" class="blob-num js-line-number" data-line-number="401"></td>
        <td id="LC401" class="blob-code blob-code-inner js-file-line">											<span class="pl-k">if</span> nowDateTime.Hour <span class="pl-k">&gt;</span> <span class="pl-c1">6</span> <span class="pl-k">and</span> nowDateTime.Hour <span class="pl-k">&lt;=</span> <span class="pl-c1">18</span>:</td>
      </tr>
      <tr>
        <td id="L402" class="blob-num js-line-number" data-line-number="402"></td>
        <td id="LC402" class="blob-code blob-code-inner js-file-line">												path <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">&quot;</span>Assets<span class="pl-cce">\\</span>Cloud-Rain-Sun.png<span class="pl-pds">&quot;</span></span></td>
      </tr>
      <tr>
        <td id="L403" class="blob-num js-line-number" data-line-number="403"></td>
        <td id="LC403" class="blob-code blob-code-inner js-file-line">																</td>
      </tr>
      <tr>
        <td id="L404" class="blob-num js-line-number" data-line-number="404"></td>
        <td id="LC404" class="blob-code blob-code-inner js-file-line">											<span class="pl-k">else</span>:</td>
      </tr>
      <tr>
        <td id="L405" class="blob-num js-line-number" data-line-number="405"></td>
        <td id="LC405" class="blob-code blob-code-inner js-file-line">												path <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">&quot;</span>Assets<span class="pl-cce">\\</span>Cloud-Rain-Moon.png<span class="pl-pds">&quot;</span></span></td>
      </tr>
      <tr>
        <td id="L406" class="blob-num js-line-number" data-line-number="406"></td>
        <td id="LC406" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L407" class="blob-num js-line-number" data-line-number="407"></td>
        <td id="LC407" class="blob-code blob-code-inner js-file-line">										<span class="pl-k">else</span>:</td>
      </tr>
      <tr>
        <td id="L408" class="blob-num js-line-number" data-line-number="408"></td>
        <td id="LC408" class="blob-code blob-code-inner js-file-line">											path <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">&quot;</span>Assets<span class="pl-cce">\\</span>Cloud-Rain.png<span class="pl-pds">&quot;</span></span></td>
      </tr>
      <tr>
        <td id="L409" class="blob-num js-line-number" data-line-number="409"></td>
        <td id="LC409" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L410" class="blob-num js-line-number" data-line-number="410"></td>
        <td id="LC410" class="blob-code blob-code-inner js-file-line">										weatherConditionList.Add(<span class="pl-s"><span class="pl-pds">&quot;</span>Rain<span class="pl-pds">&quot;</span></span>)</td>
      </tr>
      <tr>
        <td id="L411" class="blob-num js-line-number" data-line-number="411"></td>
        <td id="LC411" class="blob-code blob-code-inner js-file-line">														</td>
      </tr>
      <tr>
        <td id="L412" class="blob-num js-line-number" data-line-number="412"></td>
        <td id="LC412" class="blob-code blob-code-inner js-file-line">									<span class="pl-k">elif</span>  digit <span class="pl-k">==</span> <span class="pl-c1">6</span>:</td>
      </tr>
      <tr>
        <td id="L413" class="blob-num js-line-number" data-line-number="413"></td>
        <td id="LC413" class="blob-code blob-code-inner js-file-line">										path <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">&quot;</span>Assets<span class="pl-cce">\\</span>Cloud-Snow.png<span class="pl-pds">&quot;</span></span></td>
      </tr>
      <tr>
        <td id="L414" class="blob-num js-line-number" data-line-number="414"></td>
        <td id="LC414" class="blob-code blob-code-inner js-file-line">										weatherConditionList.Add(<span class="pl-s"><span class="pl-pds">&quot;</span>Snow<span class="pl-pds">&quot;</span></span>)</td>
      </tr>
      <tr>
        <td id="L415" class="blob-num js-line-number" data-line-number="415"></td>
        <td id="LC415" class="blob-code blob-code-inner js-file-line">														</td>
      </tr>
      <tr>
        <td id="L416" class="blob-num js-line-number" data-line-number="416"></td>
        <td id="LC416" class="blob-code blob-code-inner js-file-line">									<span class="pl-k">elif</span>  digit <span class="pl-k">==</span> <span class="pl-c1">7</span>:</td>
      </tr>
      <tr>
        <td id="L417" class="blob-num js-line-number" data-line-number="417"></td>
        <td id="LC417" class="blob-code blob-code-inner js-file-line">										path <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">&quot;</span>Assets<span class="pl-cce">\\</span>Cloud-Fog.png<span class="pl-pds">&quot;</span></span></td>
      </tr>
      <tr>
        <td id="L418" class="blob-num js-line-number" data-line-number="418"></td>
        <td id="LC418" class="blob-code blob-code-inner js-file-line">															</td>
      </tr>
      <tr>
        <td id="L419" class="blob-num js-line-number" data-line-number="419"></td>
        <td id="LC419" class="blob-code blob-code-inner js-file-line">										<span class="pl-k">if</span> Convert.ToInt32(<span class="pl-c1">id</span>) <span class="pl-k">==</span> <span class="pl-c1">701</span>:</td>
      </tr>
      <tr>
        <td id="L420" class="blob-num js-line-number" data-line-number="420"></td>
        <td id="LC420" class="blob-code blob-code-inner js-file-line">											weatherConditionList.Add(<span class="pl-s"><span class="pl-pds">&quot;</span>Mist<span class="pl-pds">&quot;</span></span>)</td>
      </tr>
      <tr>
        <td id="L421" class="blob-num js-line-number" data-line-number="421"></td>
        <td id="LC421" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L422" class="blob-num js-line-number" data-line-number="422"></td>
        <td id="LC422" class="blob-code blob-code-inner js-file-line">										<span class="pl-k">elif</span> Convert.ToInt32(<span class="pl-c1">id</span>) <span class="pl-k">==</span> <span class="pl-c1">711</span>:</td>
      </tr>
      <tr>
        <td id="L423" class="blob-num js-line-number" data-line-number="423"></td>
        <td id="LC423" class="blob-code blob-code-inner js-file-line">											weatherConditionList.Add(<span class="pl-s"><span class="pl-pds">&quot;</span>Smoke<span class="pl-pds">&quot;</span></span>)</td>
      </tr>
      <tr>
        <td id="L424" class="blob-num js-line-number" data-line-number="424"></td>
        <td id="LC424" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L425" class="blob-num js-line-number" data-line-number="425"></td>
        <td id="LC425" class="blob-code blob-code-inner js-file-line">										<span class="pl-k">elif</span> Convert.ToInt32(<span class="pl-c1">id</span>) <span class="pl-k">==</span> <span class="pl-c1">721</span>:</td>
      </tr>
      <tr>
        <td id="L426" class="blob-num js-line-number" data-line-number="426"></td>
        <td id="LC426" class="blob-code blob-code-inner js-file-line">											weatherConditionList.Add(<span class="pl-s"><span class="pl-pds">&quot;</span>Haze<span class="pl-pds">&quot;</span></span>)</td>
      </tr>
      <tr>
        <td id="L427" class="blob-num js-line-number" data-line-number="427"></td>
        <td id="LC427" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L428" class="blob-num js-line-number" data-line-number="428"></td>
        <td id="LC428" class="blob-code blob-code-inner js-file-line">										<span class="pl-k">elif</span> Convert.ToInt32(<span class="pl-c1">id</span>) <span class="pl-k">==</span> <span class="pl-c1">731</span>:</td>
      </tr>
      <tr>
        <td id="L429" class="blob-num js-line-number" data-line-number="429"></td>
        <td id="LC429" class="blob-code blob-code-inner js-file-line">											weatherConditionList.Add(<span class="pl-s"><span class="pl-pds">&quot;</span>Dust<span class="pl-pds">&quot;</span></span>)</td>
      </tr>
      <tr>
        <td id="L430" class="blob-num js-line-number" data-line-number="430"></td>
        <td id="LC430" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L431" class="blob-num js-line-number" data-line-number="431"></td>
        <td id="LC431" class="blob-code blob-code-inner js-file-line">										<span class="pl-k">elif</span> Convert.ToInt32(<span class="pl-c1">id</span>) <span class="pl-k">==</span> <span class="pl-c1">741</span>:</td>
      </tr>
      <tr>
        <td id="L432" class="blob-num js-line-number" data-line-number="432"></td>
        <td id="LC432" class="blob-code blob-code-inner js-file-line">											weatherConditionList.Add(<span class="pl-s"><span class="pl-pds">&quot;</span>Fog<span class="pl-pds">&quot;</span></span>)</td>
      </tr>
      <tr>
        <td id="L433" class="blob-num js-line-number" data-line-number="433"></td>
        <td id="LC433" class="blob-code blob-code-inner js-file-line">														</td>
      </tr>
      <tr>
        <td id="L434" class="blob-num js-line-number" data-line-number="434"></td>
        <td id="LC434" class="blob-code blob-code-inner js-file-line">									<span class="pl-k">elif</span>  digit <span class="pl-k">==</span> <span class="pl-c1">8</span>:</td>
      </tr>
      <tr>
        <td id="L435" class="blob-num js-line-number" data-line-number="435"></td>
        <td id="LC435" class="blob-code blob-code-inner js-file-line">										<span class="pl-k">if</span> Convert.ToInt32(<span class="pl-c1">id</span>) <span class="pl-k">==</span> <span class="pl-c1">800</span>:</td>
      </tr>
      <tr>
        <td id="L436" class="blob-num js-line-number" data-line-number="436"></td>
        <td id="LC436" class="blob-code blob-code-inner js-file-line">											<span class="pl-k">if</span> nowDateTime.Hour <span class="pl-k">&gt;</span> <span class="pl-c1">6</span> <span class="pl-k">and</span> nowDateTime.Hour <span class="pl-k">&lt;=</span> <span class="pl-c1">18</span>:</td>
      </tr>
      <tr>
        <td id="L437" class="blob-num js-line-number" data-line-number="437"></td>
        <td id="LC437" class="blob-code blob-code-inner js-file-line">												path <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">&quot;</span>Assets<span class="pl-cce">\\</span>Sun.png<span class="pl-pds">&quot;</span></span></td>
      </tr>
      <tr>
        <td id="L438" class="blob-num js-line-number" data-line-number="438"></td>
        <td id="LC438" class="blob-code blob-code-inner js-file-line">												weatherConditionList.Add(<span class="pl-s"><span class="pl-pds">&quot;</span>Sunny<span class="pl-pds">&quot;</span></span>)</td>
      </tr>
      <tr>
        <td id="L439" class="blob-num js-line-number" data-line-number="439"></td>
        <td id="LC439" class="blob-code blob-code-inner js-file-line">																</td>
      </tr>
      <tr>
        <td id="L440" class="blob-num js-line-number" data-line-number="440"></td>
        <td id="LC440" class="blob-code blob-code-inner js-file-line">											<span class="pl-k">else</span>:</td>
      </tr>
      <tr>
        <td id="L441" class="blob-num js-line-number" data-line-number="441"></td>
        <td id="LC441" class="blob-code blob-code-inner js-file-line">												path <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">&quot;</span>Assets<span class="pl-cce">\\</span>Moon.png<span class="pl-pds">&quot;</span></span></td>
      </tr>
      <tr>
        <td id="L442" class="blob-num js-line-number" data-line-number="442"></td>
        <td id="LC442" class="blob-code blob-code-inner js-file-line">												weatherConditionList.Add(<span class="pl-s"><span class="pl-pds">&quot;</span>Clear<span class="pl-pds">&quot;</span></span>)</td>
      </tr>
      <tr>
        <td id="L443" class="blob-num js-line-number" data-line-number="443"></td>
        <td id="LC443" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L444" class="blob-num js-line-number" data-line-number="444"></td>
        <td id="LC444" class="blob-code blob-code-inner js-file-line">										<span class="pl-k">elif</span> Convert.ToInt32(<span class="pl-c1">id</span>) <span class="pl-k">&gt;=</span> <span class="pl-c1">801</span> <span class="pl-k">and</span> Convert.ToInt32(<span class="pl-c1">id</span>) <span class="pl-k">&lt;=</span> <span class="pl-c1">803</span>:</td>
      </tr>
      <tr>
        <td id="L445" class="blob-num js-line-number" data-line-number="445"></td>
        <td id="LC445" class="blob-code blob-code-inner js-file-line">											<span class="pl-k">if</span> nowDateTime.Hour <span class="pl-k">&gt;</span> <span class="pl-c1">6</span> <span class="pl-k">and</span> nowDateTime.Hour <span class="pl-k">&lt;=</span> <span class="pl-c1">18</span>:</td>
      </tr>
      <tr>
        <td id="L446" class="blob-num js-line-number" data-line-number="446"></td>
        <td id="LC446" class="blob-code blob-code-inner js-file-line">												path <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">&quot;</span>Assets<span class="pl-cce">\\</span>Cloud-Sun.png<span class="pl-pds">&quot;</span></span></td>
      </tr>
      <tr>
        <td id="L447" class="blob-num js-line-number" data-line-number="447"></td>
        <td id="LC447" class="blob-code blob-code-inner js-file-line">																</td>
      </tr>
      <tr>
        <td id="L448" class="blob-num js-line-number" data-line-number="448"></td>
        <td id="LC448" class="blob-code blob-code-inner js-file-line">											<span class="pl-k">else</span>:</td>
      </tr>
      <tr>
        <td id="L449" class="blob-num js-line-number" data-line-number="449"></td>
        <td id="LC449" class="blob-code blob-code-inner js-file-line">												path <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">&quot;</span>Assets<span class="pl-cce">\\</span>Cloud-Moon.png<span class="pl-pds">&quot;</span></span></td>
      </tr>
      <tr>
        <td id="L450" class="blob-num js-line-number" data-line-number="450"></td>
        <td id="LC450" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L451" class="blob-num js-line-number" data-line-number="451"></td>
        <td id="LC451" class="blob-code blob-code-inner js-file-line">											weatherConditionList.Add(<span class="pl-s"><span class="pl-pds">&quot;</span>Cloudy<span class="pl-pds">&quot;</span></span>)</td>
      </tr>
      <tr>
        <td id="L452" class="blob-num js-line-number" data-line-number="452"></td>
        <td id="LC452" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L453" class="blob-num js-line-number" data-line-number="453"></td>
        <td id="LC453" class="blob-code blob-code-inner js-file-line">										<span class="pl-k">elif</span> Convert.ToInt32(<span class="pl-c1">id</span>) <span class="pl-k">==</span> <span class="pl-c1">804</span>:</td>
      </tr>
      <tr>
        <td id="L454" class="blob-num js-line-number" data-line-number="454"></td>
        <td id="LC454" class="blob-code blob-code-inner js-file-line">											path <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">&quot;</span>Assets<span class="pl-cce">\\</span>Cloud.png<span class="pl-pds">&quot;</span></span></td>
      </tr>
      <tr>
        <td id="L455" class="blob-num js-line-number" data-line-number="455"></td>
        <td id="LC455" class="blob-code blob-code-inner js-file-line">											weatherConditionList.Add(<span class="pl-s"><span class="pl-pds">&quot;</span>Overcast<span class="pl-pds">&quot;</span></span>)</td>
      </tr>
      <tr>
        <td id="L456" class="blob-num js-line-number" data-line-number="456"></td>
        <td id="LC456" class="blob-code blob-code-inner js-file-line">															</td>
      </tr>
      <tr>
        <td id="L457" class="blob-num js-line-number" data-line-number="457"></td>
        <td id="LC457" class="blob-code blob-code-inner js-file-line">									<span class="pl-k">else</span>:</td>
      </tr>
      <tr>
        <td id="L458" class="blob-num js-line-number" data-line-number="458"></td>
        <td id="LC458" class="blob-code blob-code-inner js-file-line">										<span class="pl-k">if</span> Convert.ToInt32(<span class="pl-c1">id</span>) <span class="pl-k">==</span> <span class="pl-c1">900</span>:</td>
      </tr>
      <tr>
        <td id="L459" class="blob-num js-line-number" data-line-number="459"></td>
        <td id="LC459" class="blob-code blob-code-inner js-file-line">											path <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">&quot;</span>Assets<span class="pl-cce">\\</span>Tornado.png<span class="pl-pds">&quot;</span></span></td>
      </tr>
      <tr>
        <td id="L460" class="blob-num js-line-number" data-line-number="460"></td>
        <td id="LC460" class="blob-code blob-code-inner js-file-line">											weatherConditionList.Add(<span class="pl-s"><span class="pl-pds">&quot;</span>Tornado<span class="pl-pds">&quot;</span></span>)</td>
      </tr>
      <tr>
        <td id="L461" class="blob-num js-line-number" data-line-number="461"></td>
        <td id="LC461" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L462" class="blob-num js-line-number" data-line-number="462"></td>
        <td id="LC462" class="blob-code blob-code-inner js-file-line">										<span class="pl-k">elif</span> Convert.ToInt32(<span class="pl-c1">id</span>) <span class="pl-k">==</span> <span class="pl-c1">905</span>:</td>
      </tr>
      <tr>
        <td id="L463" class="blob-num js-line-number" data-line-number="463"></td>
        <td id="LC463" class="blob-code blob-code-inner js-file-line">											path <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">&quot;</span>Assets<span class="pl-cce">\\</span>Wind.png<span class="pl-pds">&quot;</span></span></td>
      </tr>
      <tr>
        <td id="L464" class="blob-num js-line-number" data-line-number="464"></td>
        <td id="LC464" class="blob-code blob-code-inner js-file-line">											weatherConditionList.Add(<span class="pl-s"><span class="pl-pds">&quot;</span>Windy<span class="pl-pds">&quot;</span></span>)</td>
      </tr>
      <tr>
        <td id="L465" class="blob-num js-line-number" data-line-number="465"></td>
        <td id="LC465" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L466" class="blob-num js-line-number" data-line-number="466"></td>
        <td id="LC466" class="blob-code blob-code-inner js-file-line">										<span class="pl-k">elif</span> Convert.ToInt32(<span class="pl-c1">id</span>) <span class="pl-k">==</span> <span class="pl-c1">906</span>:</td>
      </tr>
      <tr>
        <td id="L467" class="blob-num js-line-number" data-line-number="467"></td>
        <td id="LC467" class="blob-code blob-code-inner js-file-line">											path <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">&quot;</span>Assets<span class="pl-cce">\\</span>Cloud-Hail.png<span class="pl-pds">&quot;</span></span></td>
      </tr>
      <tr>
        <td id="L468" class="blob-num js-line-number" data-line-number="468"></td>
        <td id="LC468" class="blob-code blob-code-inner js-file-line">											weatherConditionList.Add(<span class="pl-s"><span class="pl-pds">&quot;</span>Hail<span class="pl-pds">&quot;</span></span>)</td>
      </tr>
      <tr>
        <td id="L469" class="blob-num js-line-number" data-line-number="469"></td>
        <td id="LC469" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L470" class="blob-num js-line-number" data-line-number="470"></td>
        <td id="LC470" class="blob-code blob-code-inner js-file-line">									<span class="pl-k">if</span> path <span class="pl-k">is</span> <span class="pl-k">not</span> <span class="pl-c1">None</span> <span class="pl-k">and</span> weatherPathHashSet.Contains(path) <span class="pl-k">==</span> <span class="pl-c1">False</span>:</td>
      </tr>
      <tr>
        <td id="L471" class="blob-num js-line-number" data-line-number="471"></td>
        <td id="LC471" class="blob-code blob-code-inner js-file-line">										fs <span class="pl-k">=</span> <span class="pl-c1">None</span></td>
      </tr>
      <tr>
        <td id="L472" class="blob-num js-line-number" data-line-number="472"></td>
        <td id="LC472" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L473" class="blob-num js-line-number" data-line-number="473"></td>
        <td id="LC473" class="blob-code blob-code-inner js-file-line">										<span class="pl-k">try</span>:</td>
      </tr>
      <tr>
        <td id="L474" class="blob-num js-line-number" data-line-number="474"></td>
        <td id="LC474" class="blob-code blob-code-inner js-file-line">											fs <span class="pl-k">=</span> FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read)</td>
      </tr>
      <tr>
        <td id="L475" class="blob-num js-line-number" data-line-number="475"></td>
        <td id="LC475" class="blob-code blob-code-inner js-file-line">											ms <span class="pl-k">=</span> MemoryStream()</td>
      </tr>
      <tr>
        <td id="L476" class="blob-num js-line-number" data-line-number="476"></td>
        <td id="LC476" class="blob-code blob-code-inner js-file-line">											<span class="pl-c1">buffer</span> <span class="pl-k">=</span> Array.CreateInstance(Byte, fs.Length)</td>
      </tr>
      <tr>
        <td id="L477" class="blob-num js-line-number" data-line-number="477"></td>
        <td id="LC477" class="blob-code blob-code-inner js-file-line">											bytesRead <span class="pl-k">=</span> fs.Read(<span class="pl-c1">buffer</span>, <span class="pl-c1">0</span>, <span class="pl-c1">buffer</span>.Length)</td>
      </tr>
      <tr>
        <td id="L478" class="blob-num js-line-number" data-line-number="478"></td>
        <td id="LC478" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L479" class="blob-num js-line-number" data-line-number="479"></td>
        <td id="LC479" class="blob-code blob-code-inner js-file-line">											<span class="pl-k">while</span> bytesRead <span class="pl-k">&gt;</span> <span class="pl-c1">0</span>:</td>
      </tr>
      <tr>
        <td id="L480" class="blob-num js-line-number" data-line-number="480"></td>
        <td id="LC480" class="blob-code blob-code-inner js-file-line">												ms.Write(<span class="pl-c1">buffer</span>, <span class="pl-c1">0</span>, bytesRead)</td>
      </tr>
      <tr>
        <td id="L481" class="blob-num js-line-number" data-line-number="481"></td>
        <td id="LC481" class="blob-code blob-code-inner js-file-line">												bytesRead <span class="pl-k">=</span> fs.Read(<span class="pl-c1">buffer</span>, <span class="pl-c1">0</span>, <span class="pl-c1">buffer</span>.Length)</td>
      </tr>
      <tr>
        <td id="L482" class="blob-num js-line-number" data-line-number="482"></td>
        <td id="LC482" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L483" class="blob-num js-line-number" data-line-number="483"></td>
        <td id="LC483" class="blob-code blob-code-inner js-file-line">											ms.Seek(<span class="pl-c1">0</span>, SeekOrigin.Begin)</td>
      </tr>
      <tr>
        <td id="L484" class="blob-num js-line-number" data-line-number="484"></td>
        <td id="LC484" class="blob-code blob-code-inner js-file-line">											weatherStreamList.Add(ms)</td>
      </tr>
      <tr>
        <td id="L485" class="blob-num js-line-number" data-line-number="485"></td>
        <td id="LC485" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L486" class="blob-num js-line-number" data-line-number="486"></td>
        <td id="LC486" class="blob-code blob-code-inner js-file-line">										<span class="pl-k">finally</span>:</td>
      </tr>
      <tr>
        <td id="L487" class="blob-num js-line-number" data-line-number="487"></td>
        <td id="LC487" class="blob-code blob-code-inner js-file-line">											<span class="pl-k">if</span> fs <span class="pl-k">is</span> <span class="pl-k">not</span> <span class="pl-c1">None</span>:</td>
      </tr>
      <tr>
        <td id="L488" class="blob-num js-line-number" data-line-number="488"></td>
        <td id="LC488" class="blob-code blob-code-inner js-file-line">												fs.Close()</td>
      </tr>
      <tr>
        <td id="L489" class="blob-num js-line-number" data-line-number="489"></td>
        <td id="LC489" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L490" class="blob-num js-line-number" data-line-number="490"></td>
        <td id="LC490" class="blob-code blob-code-inner js-file-line">										weatherPathHashSet.Add(path)</td>
      </tr>
      <tr>
        <td id="L491" class="blob-num js-line-number" data-line-number="491"></td>
        <td id="LC491" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L492" class="blob-num js-line-number" data-line-number="492"></td>
        <td id="LC492" class="blob-code blob-code-inner js-file-line">				<span class="pl-k">finally</span>:</td>
      </tr>
      <tr>
        <td id="L493" class="blob-num js-line-number" data-line-number="493"></td>
        <td id="LC493" class="blob-code blob-code-inner js-file-line">					<span class="pl-k">if</span> streamReader <span class="pl-k">is</span> <span class="pl-k">not</span> <span class="pl-c1">None</span>:</td>
      </tr>
      <tr>
        <td id="L494" class="blob-num js-line-number" data-line-number="494"></td>
        <td id="LC494" class="blob-code blob-code-inner js-file-line">						streamReader.Close()</td>
      </tr>
      <tr>
        <td id="L495" class="blob-num js-line-number" data-line-number="495"></td>
        <td id="LC495" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L496" class="blob-num js-line-number" data-line-number="496"></td>
        <td id="LC496" class="blob-code blob-code-inner js-file-line">					<span class="pl-k">if</span> stream <span class="pl-k">is</span> <span class="pl-k">not</span> <span class="pl-c1">None</span>:</td>
      </tr>
      <tr>
        <td id="L497" class="blob-num js-line-number" data-line-number="497"></td>
        <td id="LC497" class="blob-code blob-code-inner js-file-line">						stream.Close()</td>
      </tr>
      <tr>
        <td id="L498" class="blob-num js-line-number" data-line-number="498"></td>
        <td id="LC498" class="blob-code blob-code-inner js-file-line">			</td>
      </tr>
      <tr>
        <td id="L499" class="blob-num js-line-number" data-line-number="499"></td>
        <td id="LC499" class="blob-code blob-code-inner js-file-line">					<span class="pl-k">if</span> response <span class="pl-k">is</span> <span class="pl-k">not</span> <span class="pl-c1">None</span>:</td>
      </tr>
      <tr>
        <td id="L500" class="blob-num js-line-number" data-line-number="500"></td>
        <td id="LC500" class="blob-code blob-code-inner js-file-line">						response.Close()</td>
      </tr>
      <tr>
        <td id="L501" class="blob-num js-line-number" data-line-number="501"></td>
        <td id="LC501" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L502" class="blob-num js-line-number" data-line-number="502"></td>
        <td id="LC502" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">except</span> <span class="pl-c1">Exception</span>, e:</td>
      </tr>
      <tr>
        <td id="L503" class="blob-num js-line-number" data-line-number="503"></td>
        <td id="LC503" class="blob-code blob-code-inner js-file-line">				Trace.WriteLine(e.clsException.Message)</td>
      </tr>
      <tr>
        <td id="L504" class="blob-num js-line-number" data-line-number="504"></td>
        <td id="LC504" class="blob-code blob-code-inner js-file-line">				Trace.WriteLine(e.clsException.StackTrace)</td>
      </tr>
      <tr>
        <td id="L505" class="blob-num js-line-number" data-line-number="505"></td>
        <td id="LC505" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L506" class="blob-num js-line-number" data-line-number="506"></td>
        <td id="LC506" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">return</span> KeyValuePair[List[String], KeyValuePair[Double, KeyValuePair[KeyValuePair[Double, Double], KeyValuePair[List[Double], List[MemoryStream]]]]](weatherConditionList, KeyValuePair[Double, KeyValuePair[KeyValuePair[Double, Double], KeyValuePair[List[Double], List[MemoryStream]]]](temp, KeyValuePair[KeyValuePair[Double, Double], KeyValuePair[List[Double], List[MemoryStream]]](KeyValuePair[Double, Double](windSpeed, windDeg), KeyValuePair[List[Double], List[MemoryStream]](weatherIdList, weatherStreamList))))</td>
      </tr>
      <tr>
        <td id="L507" class="blob-num js-line-number" data-line-number="507"></td>
        <td id="LC507" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L508" class="blob-num js-line-number" data-line-number="508"></td>
        <td id="LC508" class="blob-code blob-code-inner js-file-line">	<span class="pl-k">def</span> <span class="pl-en">onCompleted</span>(<span class="pl-smi">task</span>):</td>
      </tr>
      <tr>
        <td id="L509" class="blob-num js-line-number" data-line-number="509"></td>
        <td id="LC509" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">global</span> idList</td>
      </tr>
      <tr>
        <td id="L510" class="blob-num js-line-number" data-line-number="510"></td>
        <td id="LC510" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L511" class="blob-num js-line-number" data-line-number="511"></td>
        <td id="LC511" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">if</span> task.Result.Key.Count <span class="pl-k">&gt;</span> <span class="pl-c1">0</span>:</td>
      </tr>
      <tr>
        <td id="L512" class="blob-num js-line-number" data-line-number="512"></td>
        <td id="LC512" class="blob-code blob-code-inner js-file-line">			sequenceList <span class="pl-k">=</span> List[Sequence]()</td>
      </tr>
      <tr>
        <td id="L513" class="blob-num js-line-number" data-line-number="513"></td>
        <td id="LC513" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L514" class="blob-num js-line-number" data-line-number="514"></td>
        <td id="LC514" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">for</span> sequence <span class="pl-k">in</span> Script.Instance.Sequences:</td>
      </tr>
      <tr>
        <td id="L515" class="blob-num js-line-number" data-line-number="515"></td>
        <td id="LC515" class="blob-code blob-code-inner js-file-line">				<span class="pl-k">if</span> sequence.Name.Equals(<span class="pl-s"><span class="pl-pds">&quot;</span>Weather<span class="pl-pds">&quot;</span></span>):</td>
      </tr>
      <tr>
        <td id="L516" class="blob-num js-line-number" data-line-number="516"></td>
        <td id="LC516" class="blob-code blob-code-inner js-file-line">					sequenceList.Add(sequence)</td>
      </tr>
      <tr>
        <td id="L517" class="blob-num js-line-number" data-line-number="517"></td>
        <td id="LC517" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L518" class="blob-num js-line-number" data-line-number="518"></td>
        <td id="LC518" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">for</span> s <span class="pl-k">in</span> task.Result.Key:</td>
      </tr>
      <tr>
        <td id="L519" class="blob-num js-line-number" data-line-number="519"></td>
        <td id="LC519" class="blob-code blob-code-inner js-file-line">				Script.Instance.TryEnqueue(Script.Instance.Prepare(sequenceList, s))</td>
      </tr>
      <tr>
        <td id="L520" class="blob-num js-line-number" data-line-number="520"></td>
        <td id="LC520" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L521" class="blob-num js-line-number" data-line-number="521"></td>
        <td id="LC521" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">if</span> Application.Current.MainWindow.IsVisible <span class="pl-k">and</span> task.Result.Value.Value.Value.Value.Count <span class="pl-k">&gt;</span> <span class="pl-c1">0</span> <span class="pl-k">and</span> <span class="pl-k">not</span> Enumerable.SequenceEqual[Double](idList, task.Result.Value.Value.Value.Key):</td>
      </tr>
      <tr>
        <td id="L522" class="blob-num js-line-number" data-line-number="522"></td>
        <td id="LC522" class="blob-code blob-code-inner js-file-line">			width <span class="pl-k">=</span> <span class="pl-c1">128</span></td>
      </tr>
      <tr>
        <td id="L523" class="blob-num js-line-number" data-line-number="523"></td>
        <td id="LC523" class="blob-code blob-code-inner js-file-line">			height <span class="pl-k">=</span> <span class="pl-c1">128</span></td>
      </tr>
      <tr>
        <td id="L524" class="blob-num js-line-number" data-line-number="524"></td>
        <td id="LC524" class="blob-code blob-code-inner js-file-line">			<span class="pl-c1">max</span> <span class="pl-k">=</span> <span class="pl-c1">4</span></td>
      </tr>
      <tr>
        <td id="L525" class="blob-num js-line-number" data-line-number="525"></td>
        <td id="LC525" class="blob-code blob-code-inner js-file-line">			window <span class="pl-k">=</span> Window()</td>
      </tr>
      <tr>
        <td id="L526" class="blob-num js-line-number" data-line-number="526"></td>
        <td id="LC526" class="blob-code blob-code-inner js-file-line">			contentControl <span class="pl-k">=</span> ContentControl()</td>
      </tr>
      <tr>
        <td id="L527" class="blob-num js-line-number" data-line-number="527"></td>
        <td id="LC527" class="blob-code blob-code-inner js-file-line">			grid <span class="pl-k">=</span> Grid()</td>
      </tr>
      <tr>
        <td id="L528" class="blob-num js-line-number" data-line-number="528"></td>
        <td id="LC528" class="blob-code blob-code-inner js-file-line">			storyboard <span class="pl-k">=</span> Storyboard()</td>
      </tr>
      <tr>
        <td id="L529" class="blob-num js-line-number" data-line-number="529"></td>
        <td id="LC529" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L530" class="blob-num js-line-number" data-line-number="530"></td>
        <td id="LC530" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">def</span> <span class="pl-en">onCurrentStateInvalidated</span>(<span class="pl-smi">sender</span>, <span class="pl-smi">args</span>):</td>
      </tr>
      <tr>
        <td id="L531" class="blob-num js-line-number" data-line-number="531"></td>
        <td id="LC531" class="blob-code blob-code-inner js-file-line">				<span class="pl-k">if</span> sender.CurrentState <span class="pl-k">==</span> ClockState.Filling:</td>
      </tr>
      <tr>
        <td id="L532" class="blob-num js-line-number" data-line-number="532"></td>
        <td id="LC532" class="blob-code blob-code-inner js-file-line">					window.Close()</td>
      </tr>
      <tr>
        <td id="L533" class="blob-num js-line-number" data-line-number="533"></td>
        <td id="LC533" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L534" class="blob-num js-line-number" data-line-number="534"></td>
        <td id="LC534" class="blob-code blob-code-inner js-file-line">			storyboard.CurrentStateInvalidated <span class="pl-k">+=</span> onCurrentStateInvalidated</td>
      </tr>
      <tr>
        <td id="L535" class="blob-num js-line-number" data-line-number="535"></td>
        <td id="LC535" class="blob-code blob-code-inner js-file-line">		</td>
      </tr>
      <tr>
        <td id="L536" class="blob-num js-line-number" data-line-number="536"></td>
        <td id="LC536" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">def</span> <span class="pl-en">onLoaded</span>(<span class="pl-smi">sender</span>, <span class="pl-smi">args</span>):</td>
      </tr>
      <tr>
        <td id="L537" class="blob-num js-line-number" data-line-number="537"></td>
        <td id="LC537" class="blob-code blob-code-inner js-file-line">				time <span class="pl-k">=</span> <span class="pl-c1">0</span></td>
      </tr>
      <tr>
        <td id="L538" class="blob-num js-line-number" data-line-number="538"></td>
        <td id="LC538" class="blob-code blob-code-inner js-file-line">				speed <span class="pl-k">=</span> task.Result.Value.Value.Key.Key <span class="pl-k">*</span> <span class="pl-c1">1000</span> <span class="pl-k">/</span> <span class="pl-c1">60</span> <span class="pl-k">/</span> <span class="pl-c1">60</span></td>
      </tr>
      <tr>
        <td id="L539" class="blob-num js-line-number" data-line-number="539"></td>
        <td id="LC539" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L540" class="blob-num js-line-number" data-line-number="540"></td>
        <td id="LC540" class="blob-code blob-code-inner js-file-line">				contentControl.Width <span class="pl-k">=</span> contentControl.ActualWidth <span class="pl-k">*</span> <span class="pl-c1">1.5</span> <span class="pl-k">if</span> contentControl.ActualWidth <span class="pl-k">&gt;</span> contentControl.ActualHeight <span class="pl-k">else</span> contentControl.ActualHeight <span class="pl-k">*</span> <span class="pl-c1">1.5</span> </td>
      </tr>
      <tr>
        <td id="L541" class="blob-num js-line-number" data-line-number="541"></td>
        <td id="LC541" class="blob-code blob-code-inner js-file-line">				contentControl.Height <span class="pl-k">=</span> contentControl.ActualWidth <span class="pl-k">*</span> <span class="pl-c1">1.5</span> <span class="pl-k">if</span> contentControl.ActualWidth <span class="pl-k">&gt;</span> contentControl.ActualHeight <span class="pl-k">else</span> contentControl.ActualHeight <span class="pl-k">*</span> <span class="pl-c1">1.5</span></td>
      </tr>
      <tr>
        <td id="L542" class="blob-num js-line-number" data-line-number="542"></td>
        <td id="LC542" class="blob-code blob-code-inner js-file-line">				contentControl.RenderTransform.CenterX <span class="pl-k">=</span> contentControl.Width <span class="pl-k">/</span> <span class="pl-c1">2</span></td>
      </tr>
      <tr>
        <td id="L543" class="blob-num js-line-number" data-line-number="543"></td>
        <td id="LC543" class="blob-code blob-code-inner js-file-line">				contentControl.RenderTransform.CenterY <span class="pl-k">=</span> contentControl.Height <span class="pl-k">/</span> <span class="pl-c1">2</span></td>
      </tr>
      <tr>
        <td id="L544" class="blob-num js-line-number" data-line-number="544"></td>
        <td id="LC544" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L545" class="blob-num js-line-number" data-line-number="545"></td>
        <td id="LC545" class="blob-code blob-code-inner js-file-line">				doubleAnimation1 <span class="pl-k">=</span> DoubleAnimation(contentControl.Opacity, <span class="pl-c1">1</span>, TimeSpan.FromMilliseconds(<span class="pl-c1">500</span>))</td>
      </tr>
      <tr>
        <td id="L546" class="blob-num js-line-number" data-line-number="546"></td>
        <td id="LC546" class="blob-code blob-code-inner js-file-line">				doubleAnimation2 <span class="pl-k">=</span> DoubleAnimation(<span class="pl-c1">1.5</span>, <span class="pl-c1">1</span>, TimeSpan.FromMilliseconds(<span class="pl-c1">500</span>))</td>
      </tr>
      <tr>
        <td id="L547" class="blob-num js-line-number" data-line-number="547"></td>
        <td id="LC547" class="blob-code blob-code-inner js-file-line">				doubleAnimation3 <span class="pl-k">=</span> DoubleAnimation(<span class="pl-c1">1.5</span>, <span class="pl-c1">1</span>, TimeSpan.FromMilliseconds(<span class="pl-c1">500</span>))</td>
      </tr>
      <tr>
        <td id="L548" class="blob-num js-line-number" data-line-number="548"></td>
        <td id="LC548" class="blob-code blob-code-inner js-file-line">				doubleAnimation4 <span class="pl-k">=</span> DoubleAnimation(<span class="pl-c1">1</span>, <span class="pl-c1">0</span>, TimeSpan.FromMilliseconds(<span class="pl-c1">500</span>))</td>
      </tr>
      <tr>
        <td id="L549" class="blob-num js-line-number" data-line-number="549"></td>
        <td id="LC549" class="blob-code blob-code-inner js-file-line">				doubleAnimation5 <span class="pl-k">=</span> DoubleAnimation(<span class="pl-c1">1</span>, <span class="pl-c1">1.5</span>, TimeSpan.FromMilliseconds(<span class="pl-c1">500</span>))</td>
      </tr>
      <tr>
        <td id="L550" class="blob-num js-line-number" data-line-number="550"></td>
        <td id="LC550" class="blob-code blob-code-inner js-file-line">				doubleAnimation6 <span class="pl-k">=</span> DoubleAnimation(<span class="pl-c1">1</span>, <span class="pl-c1">1.5</span>, TimeSpan.FromMilliseconds(<span class="pl-c1">500</span>))</td>
      </tr>
      <tr>
        <td id="L551" class="blob-num js-line-number" data-line-number="551"></td>
        <td id="LC551" class="blob-code blob-code-inner js-file-line">				sineEase1 <span class="pl-k">=</span> SineEase()</td>
      </tr>
      <tr>
        <td id="L552" class="blob-num js-line-number" data-line-number="552"></td>
        <td id="LC552" class="blob-code blob-code-inner js-file-line">				sineEase2 <span class="pl-k">=</span> SineEase()</td>
      </tr>
      <tr>
        <td id="L553" class="blob-num js-line-number" data-line-number="553"></td>
        <td id="LC553" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L554" class="blob-num js-line-number" data-line-number="554"></td>
        <td id="LC554" class="blob-code blob-code-inner js-file-line">				sineEase1.EasingMode <span class="pl-k">=</span> EasingMode.EaseOut</td>
      </tr>
      <tr>
        <td id="L555" class="blob-num js-line-number" data-line-number="555"></td>
        <td id="LC555" class="blob-code blob-code-inner js-file-line">				sineEase2.EasingMode <span class="pl-k">=</span> EasingMode.EaseIn</td>
      </tr>
      <tr>
        <td id="L556" class="blob-num js-line-number" data-line-number="556"></td>
        <td id="LC556" class="blob-code blob-code-inner js-file-line">				doubleAnimation1.EasingFunction <span class="pl-k">=</span> doubleAnimation2.EasingFunction <span class="pl-k">=</span> doubleAnimation3.EasingFunction <span class="pl-k">=</span> sineEase1</td>
      </tr>
      <tr>
        <td id="L557" class="blob-num js-line-number" data-line-number="557"></td>
        <td id="LC557" class="blob-code blob-code-inner js-file-line">				doubleAnimation4.EasingFunction <span class="pl-k">=</span> doubleAnimation5.EasingFunction <span class="pl-k">=</span> doubleAnimation6.EasingFunction <span class="pl-k">=</span> sineEase2</td>
      </tr>
      <tr>
        <td id="L558" class="blob-num js-line-number" data-line-number="558"></td>
        <td id="LC558" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L559" class="blob-num js-line-number" data-line-number="559"></td>
        <td id="LC559" class="blob-code blob-code-inner js-file-line">				doubleAnimation4.BeginTime <span class="pl-k">=</span> Nullable[TimeSpan](TimeSpan.FromMilliseconds((<span class="pl-c1">250</span> <span class="pl-k">*</span> (<span class="pl-c1">max</span> <span class="pl-k">-</span> <span class="pl-c1">1</span>) <span class="pl-k">*</span> <span class="pl-c1">2</span> <span class="pl-k">+</span> <span class="pl-c1">1000</span> <span class="pl-k">+</span> <span class="pl-c1">3000</span>) <span class="pl-k">*</span> task.Result.Value.Value.Value.Value.Count <span class="pl-k">-</span> <span class="pl-c1">500</span>))</td>
      </tr>
      <tr>
        <td id="L560" class="blob-num js-line-number" data-line-number="560"></td>
        <td id="LC560" class="blob-code blob-code-inner js-file-line">				doubleAnimation5.BeginTime <span class="pl-k">=</span> Nullable[TimeSpan](TimeSpan.FromMilliseconds((<span class="pl-c1">250</span> <span class="pl-k">*</span> (<span class="pl-c1">max</span> <span class="pl-k">-</span> <span class="pl-c1">1</span>) <span class="pl-k">*</span> <span class="pl-c1">2</span> <span class="pl-k">+</span> <span class="pl-c1">1000</span> <span class="pl-k">+</span> <span class="pl-c1">3000</span>) <span class="pl-k">*</span> task.Result.Value.Value.Value.Value.Count <span class="pl-k">-</span> <span class="pl-c1">500</span>))</td>
      </tr>
      <tr>
        <td id="L561" class="blob-num js-line-number" data-line-number="561"></td>
        <td id="LC561" class="blob-code blob-code-inner js-file-line">				doubleAnimation6.BeginTime <span class="pl-k">=</span> Nullable[TimeSpan](TimeSpan.FromMilliseconds((<span class="pl-c1">250</span> <span class="pl-k">*</span> (<span class="pl-c1">max</span> <span class="pl-k">-</span> <span class="pl-c1">1</span>) <span class="pl-k">*</span> <span class="pl-c1">2</span> <span class="pl-k">+</span> <span class="pl-c1">1000</span> <span class="pl-k">+</span> <span class="pl-c1">3000</span>) <span class="pl-k">*</span> task.Result.Value.Value.Value.Value.Count <span class="pl-k">-</span> <span class="pl-c1">500</span>))</td>
      </tr>
      <tr>
        <td id="L562" class="blob-num js-line-number" data-line-number="562"></td>
        <td id="LC562" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L563" class="blob-num js-line-number" data-line-number="563"></td>
        <td id="LC563" class="blob-code blob-code-inner js-file-line">				storyboard.Children.Add(doubleAnimation1)</td>
      </tr>
      <tr>
        <td id="L564" class="blob-num js-line-number" data-line-number="564"></td>
        <td id="LC564" class="blob-code blob-code-inner js-file-line">				storyboard.Children.Add(doubleAnimation2)</td>
      </tr>
      <tr>
        <td id="L565" class="blob-num js-line-number" data-line-number="565"></td>
        <td id="LC565" class="blob-code blob-code-inner js-file-line">				storyboard.Children.Add(doubleAnimation3)</td>
      </tr>
      <tr>
        <td id="L566" class="blob-num js-line-number" data-line-number="566"></td>
        <td id="LC566" class="blob-code blob-code-inner js-file-line">				storyboard.Children.Add(doubleAnimation4)</td>
      </tr>
      <tr>
        <td id="L567" class="blob-num js-line-number" data-line-number="567"></td>
        <td id="LC567" class="blob-code blob-code-inner js-file-line">				storyboard.Children.Add(doubleAnimation5)</td>
      </tr>
      <tr>
        <td id="L568" class="blob-num js-line-number" data-line-number="568"></td>
        <td id="LC568" class="blob-code blob-code-inner js-file-line">				storyboard.Children.Add(doubleAnimation6)</td>
      </tr>
      <tr>
        <td id="L569" class="blob-num js-line-number" data-line-number="569"></td>
        <td id="LC569" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L570" class="blob-num js-line-number" data-line-number="570"></td>
        <td id="LC570" class="blob-code blob-code-inner js-file-line">				Storyboard.SetTarget(doubleAnimation1, contentControl)</td>
      </tr>
      <tr>
        <td id="L571" class="blob-num js-line-number" data-line-number="571"></td>
        <td id="LC571" class="blob-code blob-code-inner js-file-line">				Storyboard.SetTarget(doubleAnimation2, contentControl)</td>
      </tr>
      <tr>
        <td id="L572" class="blob-num js-line-number" data-line-number="572"></td>
        <td id="LC572" class="blob-code blob-code-inner js-file-line">				Storyboard.SetTarget(doubleAnimation3, contentControl)</td>
      </tr>
      <tr>
        <td id="L573" class="blob-num js-line-number" data-line-number="573"></td>
        <td id="LC573" class="blob-code blob-code-inner js-file-line">				Storyboard.SetTarget(doubleAnimation4, contentControl)</td>
      </tr>
      <tr>
        <td id="L574" class="blob-num js-line-number" data-line-number="574"></td>
        <td id="LC574" class="blob-code blob-code-inner js-file-line">				Storyboard.SetTarget(doubleAnimation5, contentControl)</td>
      </tr>
      <tr>
        <td id="L575" class="blob-num js-line-number" data-line-number="575"></td>
        <td id="LC575" class="blob-code blob-code-inner js-file-line">				Storyboard.SetTarget(doubleAnimation6, contentControl)</td>
      </tr>
      <tr>
        <td id="L576" class="blob-num js-line-number" data-line-number="576"></td>
        <td id="LC576" class="blob-code blob-code-inner js-file-line">				Storyboard.SetTargetProperty(doubleAnimation1, PropertyPath(ContentControl.OpacityProperty))</td>
      </tr>
      <tr>
        <td id="L577" class="blob-num js-line-number" data-line-number="577"></td>
        <td id="LC577" class="blob-code blob-code-inner js-file-line">				Storyboard.SetTargetProperty(doubleAnimation2, PropertyPath(<span class="pl-s"><span class="pl-pds">&quot;</span>(0).(1)<span class="pl-pds">&quot;</span></span>, ContentControl.RenderTransformProperty, ScaleTransform.ScaleXProperty))</td>
      </tr>
      <tr>
        <td id="L578" class="blob-num js-line-number" data-line-number="578"></td>
        <td id="LC578" class="blob-code blob-code-inner js-file-line">				Storyboard.SetTargetProperty(doubleAnimation3, PropertyPath(<span class="pl-s"><span class="pl-pds">&quot;</span>(0).(1)<span class="pl-pds">&quot;</span></span>, ContentControl.RenderTransformProperty, ScaleTransform.ScaleYProperty))</td>
      </tr>
      <tr>
        <td id="L579" class="blob-num js-line-number" data-line-number="579"></td>
        <td id="LC579" class="blob-code blob-code-inner js-file-line">				Storyboard.SetTargetProperty(doubleAnimation4, PropertyPath(ContentControl.OpacityProperty))</td>
      </tr>
      <tr>
        <td id="L580" class="blob-num js-line-number" data-line-number="580"></td>
        <td id="LC580" class="blob-code blob-code-inner js-file-line">				Storyboard.SetTargetProperty(doubleAnimation5, PropertyPath(<span class="pl-s"><span class="pl-pds">&quot;</span>(0).(1)<span class="pl-pds">&quot;</span></span>, ContentControl.RenderTransformProperty, ScaleTransform.ScaleXProperty))</td>
      </tr>
      <tr>
        <td id="L581" class="blob-num js-line-number" data-line-number="581"></td>
        <td id="LC581" class="blob-code blob-code-inner js-file-line">				Storyboard.SetTargetProperty(doubleAnimation6, PropertyPath(<span class="pl-s"><span class="pl-pds">&quot;</span>(0).(1)<span class="pl-pds">&quot;</span></span>, ContentControl.RenderTransformProperty, ScaleTransform.ScaleYProperty))</td>
      </tr>
      <tr>
        <td id="L582" class="blob-num js-line-number" data-line-number="582"></td>
        <td id="LC582" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L583" class="blob-num js-line-number" data-line-number="583"></td>
        <td id="LC583" class="blob-code blob-code-inner js-file-line">				<span class="pl-k">for</span> element1 <span class="pl-k">in</span> grid.Children:</td>
      </tr>
      <tr>
        <td id="L584" class="blob-num js-line-number" data-line-number="584"></td>
        <td id="LC584" class="blob-code blob-code-inner js-file-line">					<span class="pl-k">for</span> element2 <span class="pl-k">in</span> element1.Children:</td>
      </tr>
      <tr>
        <td id="L585" class="blob-num js-line-number" data-line-number="585"></td>
        <td id="LC585" class="blob-code blob-code-inner js-file-line">						w <span class="pl-k">=</span> element2.Width <span class="pl-k">/</span> <span class="pl-c1">2</span> <span class="pl-k">if</span> speed <span class="pl-k">&gt;</span> <span class="pl-c1">15</span> <span class="pl-k">else</span> element2.Width <span class="pl-k">/</span> <span class="pl-c1">2</span> <span class="pl-k">*</span> speed <span class="pl-k">/</span> <span class="pl-c1">15</span>;</td>
      </tr>
      <tr>
        <td id="L586" class="blob-num js-line-number" data-line-number="586"></td>
        <td id="LC586" class="blob-code blob-code-inner js-file-line">						da1 <span class="pl-k">=</span> DoubleAnimation(element2.Opacity, <span class="pl-c1">1</span>, TimeSpan.FromMilliseconds(<span class="pl-c1">1000</span>))</td>
      </tr>
      <tr>
        <td id="L587" class="blob-num js-line-number" data-line-number="587"></td>
        <td id="LC587" class="blob-code blob-code-inner js-file-line">						da2 <span class="pl-k">=</span> DoubleAnimation(<span class="pl-k">-</span>w <span class="pl-k">if</span> Convert.ToInt32(task.Result.Value.Value.Key.Value <span class="pl-k">/</span> <span class="pl-c1">180</span>) <span class="pl-k">%</span> <span class="pl-c1">2</span> <span class="pl-k">==</span> <span class="pl-c1">0</span> <span class="pl-k">else</span> w, <span class="pl-c1">0</span>, TimeSpan.FromMilliseconds(<span class="pl-c1">1000</span>))</td>
      </tr>
      <tr>
        <td id="L588" class="blob-num js-line-number" data-line-number="588"></td>
        <td id="LC588" class="blob-code blob-code-inner js-file-line">						da3 <span class="pl-k">=</span> DoubleAnimation(<span class="pl-c1">1</span>, <span class="pl-c1">0</span>, TimeSpan.FromMilliseconds(<span class="pl-c1">1000</span>))</td>
      </tr>
      <tr>
        <td id="L589" class="blob-num js-line-number" data-line-number="589"></td>
        <td id="LC589" class="blob-code blob-code-inner js-file-line">						da4 <span class="pl-k">=</span> DoubleAnimation(<span class="pl-c1">0</span>, w <span class="pl-k">if</span> Convert.ToInt32(task.Result.Value.Value.Key.Value <span class="pl-k">/</span> <span class="pl-c1">180</span>) <span class="pl-k">%</span> <span class="pl-c1">2</span> <span class="pl-k">==</span> <span class="pl-c1">0</span> <span class="pl-k">else</span> <span class="pl-k">-</span>w, TimeSpan.FromMilliseconds(<span class="pl-c1">1000</span>))</td>
      </tr>
      <tr>
        <td id="L590" class="blob-num js-line-number" data-line-number="590"></td>
        <td id="LC590" class="blob-code blob-code-inner js-file-line">						se1 <span class="pl-k">=</span> SineEase()</td>
      </tr>
      <tr>
        <td id="L591" class="blob-num js-line-number" data-line-number="591"></td>
        <td id="LC591" class="blob-code blob-code-inner js-file-line">						se2 <span class="pl-k">=</span> SineEase()</td>
      </tr>
      <tr>
        <td id="L592" class="blob-num js-line-number" data-line-number="592"></td>
        <td id="LC592" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L593" class="blob-num js-line-number" data-line-number="593"></td>
        <td id="LC593" class="blob-code blob-code-inner js-file-line">						da1.BeginTime <span class="pl-k">=</span> Nullable[TimeSpan](TimeSpan.FromMilliseconds(time <span class="pl-k">+</span> <span class="pl-c1">250</span> <span class="pl-k">*</span> element2.Tag))</td>
      </tr>
      <tr>
        <td id="L594" class="blob-num js-line-number" data-line-number="594"></td>
        <td id="LC594" class="blob-code blob-code-inner js-file-line">						da2.BeginTime <span class="pl-k">=</span> Nullable[TimeSpan](TimeSpan.FromMilliseconds(time <span class="pl-k">+</span> <span class="pl-c1">250</span> <span class="pl-k">*</span> element2.Tag))</td>
      </tr>
      <tr>
        <td id="L595" class="blob-num js-line-number" data-line-number="595"></td>
        <td id="LC595" class="blob-code blob-code-inner js-file-line">						da3.BeginTime <span class="pl-k">=</span> Nullable[TimeSpan](TimeSpan.FromMilliseconds(time <span class="pl-k">+</span> <span class="pl-c1">250</span> <span class="pl-k">*</span> element2.Tag <span class="pl-k">*</span> <span class="pl-c1">2</span> <span class="pl-k">+</span> <span class="pl-c1">250</span> <span class="pl-k">*</span> (<span class="pl-c1">max</span> <span class="pl-k">-</span> <span class="pl-c1">1</span>) <span class="pl-k">-</span> <span class="pl-c1">250</span> <span class="pl-k">*</span> element2.Tag <span class="pl-k">+</span> <span class="pl-c1">3000</span>))</td>
      </tr>
      <tr>
        <td id="L596" class="blob-num js-line-number" data-line-number="596"></td>
        <td id="LC596" class="blob-code blob-code-inner js-file-line">						da4.BeginTime <span class="pl-k">=</span> Nullable[TimeSpan](TimeSpan.FromMilliseconds(time <span class="pl-k">+</span> <span class="pl-c1">250</span> <span class="pl-k">*</span> element2.Tag <span class="pl-k">*</span> <span class="pl-c1">2</span> <span class="pl-k">+</span> <span class="pl-c1">250</span> <span class="pl-k">*</span> (<span class="pl-c1">max</span> <span class="pl-k">-</span> <span class="pl-c1">1</span>) <span class="pl-k">-</span> <span class="pl-c1">250</span> <span class="pl-k">*</span> element2.Tag <span class="pl-k">+</span> <span class="pl-c1">3000</span>))</td>
      </tr>
      <tr>
        <td id="L597" class="blob-num js-line-number" data-line-number="597"></td>
        <td id="LC597" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L598" class="blob-num js-line-number" data-line-number="598"></td>
        <td id="LC598" class="blob-code blob-code-inner js-file-line">						se1.EasingMode <span class="pl-k">=</span> EasingMode.EaseOut</td>
      </tr>
      <tr>
        <td id="L599" class="blob-num js-line-number" data-line-number="599"></td>
        <td id="LC599" class="blob-code blob-code-inner js-file-line">						se2.EasingMode <span class="pl-k">=</span> EasingMode.EaseIn</td>
      </tr>
      <tr>
        <td id="L600" class="blob-num js-line-number" data-line-number="600"></td>
        <td id="LC600" class="blob-code blob-code-inner js-file-line">						da1.EasingFunction <span class="pl-k">=</span> da2.EasingFunction <span class="pl-k">=</span> se1</td>
      </tr>
      <tr>
        <td id="L601" class="blob-num js-line-number" data-line-number="601"></td>
        <td id="LC601" class="blob-code blob-code-inner js-file-line">						da3.EasingFunction <span class="pl-k">=</span> da4.EasingFunction <span class="pl-k">=</span> se2</td>
      </tr>
      <tr>
        <td id="L602" class="blob-num js-line-number" data-line-number="602"></td>
        <td id="LC602" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L603" class="blob-num js-line-number" data-line-number="603"></td>
        <td id="LC603" class="blob-code blob-code-inner js-file-line">						storyboard.Children.Add(da1)</td>
      </tr>
      <tr>
        <td id="L604" class="blob-num js-line-number" data-line-number="604"></td>
        <td id="LC604" class="blob-code blob-code-inner js-file-line">						storyboard.Children.Add(da2)</td>
      </tr>
      <tr>
        <td id="L605" class="blob-num js-line-number" data-line-number="605"></td>
        <td id="LC605" class="blob-code blob-code-inner js-file-line">						storyboard.Children.Add(da3)</td>
      </tr>
      <tr>
        <td id="L606" class="blob-num js-line-number" data-line-number="606"></td>
        <td id="LC606" class="blob-code blob-code-inner js-file-line">						storyboard.Children.Add(da4)</td>
      </tr>
      <tr>
        <td id="L607" class="blob-num js-line-number" data-line-number="607"></td>
        <td id="LC607" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L608" class="blob-num js-line-number" data-line-number="608"></td>
        <td id="LC608" class="blob-code blob-code-inner js-file-line">						Storyboard.SetTarget(da1, element2)</td>
      </tr>
      <tr>
        <td id="L609" class="blob-num js-line-number" data-line-number="609"></td>
        <td id="LC609" class="blob-code blob-code-inner js-file-line">						Storyboard.SetTarget(da2, element2)</td>
      </tr>
      <tr>
        <td id="L610" class="blob-num js-line-number" data-line-number="610"></td>
        <td id="LC610" class="blob-code blob-code-inner js-file-line">						Storyboard.SetTarget(da3, element2)</td>
      </tr>
      <tr>
        <td id="L611" class="blob-num js-line-number" data-line-number="611"></td>
        <td id="LC611" class="blob-code blob-code-inner js-file-line">						Storyboard.SetTarget(da4, element2)</td>
      </tr>
      <tr>
        <td id="L612" class="blob-num js-line-number" data-line-number="612"></td>
        <td id="LC612" class="blob-code blob-code-inner js-file-line">						Storyboard.SetTargetProperty(da1, PropertyPath(Rectangle.OpacityProperty))</td>
      </tr>
      <tr>
        <td id="L613" class="blob-num js-line-number" data-line-number="613"></td>
        <td id="LC613" class="blob-code blob-code-inner js-file-line">						Storyboard.SetTargetProperty(da2, PropertyPath(<span class="pl-s"><span class="pl-pds">&quot;</span>(0).(1)<span class="pl-pds">&quot;</span></span>, Rectangle.RenderTransformProperty, TranslateTransform.XProperty))</td>
      </tr>
      <tr>
        <td id="L614" class="blob-num js-line-number" data-line-number="614"></td>
        <td id="LC614" class="blob-code blob-code-inner js-file-line">						Storyboard.SetTargetProperty(da3, PropertyPath(Rectangle.OpacityProperty))</td>
      </tr>
      <tr>
        <td id="L615" class="blob-num js-line-number" data-line-number="615"></td>
        <td id="LC615" class="blob-code blob-code-inner js-file-line">						Storyboard.SetTargetProperty(da4, PropertyPath(<span class="pl-s"><span class="pl-pds">&quot;</span>(0).(1)<span class="pl-pds">&quot;</span></span>, Rectangle.RenderTransformProperty, TranslateTransform.XProperty))</td>
      </tr>
      <tr>
        <td id="L616" class="blob-num js-line-number" data-line-number="616"></td>
        <td id="LC616" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L617" class="blob-num js-line-number" data-line-number="617"></td>
        <td id="LC617" class="blob-code blob-code-inner js-file-line">					time <span class="pl-k">+=</span> <span class="pl-c1">250</span> <span class="pl-k">*</span> (<span class="pl-c1">max</span> <span class="pl-k">-</span> <span class="pl-c1">1</span>) <span class="pl-k">+</span> <span class="pl-c1">1000</span> <span class="pl-k">+</span> <span class="pl-c1">3000</span></td>
      </tr>
      <tr>
        <td id="L618" class="blob-num js-line-number" data-line-number="618"></td>
        <td id="LC618" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L619" class="blob-num js-line-number" data-line-number="619"></td>
        <td id="LC619" class="blob-code blob-code-inner js-file-line">				storyboard.Begin()</td>
      </tr>
      <tr>
        <td id="L620" class="blob-num js-line-number" data-line-number="620"></td>
        <td id="LC620" class="blob-code blob-code-inner js-file-line">			</td>
      </tr>
      <tr>
        <td id="L621" class="blob-num js-line-number" data-line-number="621"></td>
        <td id="LC621" class="blob-code blob-code-inner js-file-line">			fs <span class="pl-k">=</span> <span class="pl-c1">None</span></td>
      </tr>
      <tr>
        <td id="L622" class="blob-num js-line-number" data-line-number="622"></td>
        <td id="LC622" class="blob-code blob-code-inner js-file-line">			bi <span class="pl-k">=</span> BitmapImage()</td>
      </tr>
      <tr>
        <td id="L623" class="blob-num js-line-number" data-line-number="623"></td>
        <td id="LC623" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L624" class="blob-num js-line-number" data-line-number="624"></td>
        <td id="LC624" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">try</span>:</td>
      </tr>
      <tr>
        <td id="L625" class="blob-num js-line-number" data-line-number="625"></td>
        <td id="LC625" class="blob-code blob-code-inner js-file-line">				fs <span class="pl-k">=</span> FileStream(<span class="pl-s"><span class="pl-pds">&quot;</span>Assets<span class="pl-cce">\\</span>Background-Popup.png<span class="pl-pds">&quot;</span></span>, FileMode.Open, FileAccess.Read, FileShare.Read)</td>
      </tr>
      <tr>
        <td id="L626" class="blob-num js-line-number" data-line-number="626"></td>
        <td id="LC626" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L627" class="blob-num js-line-number" data-line-number="627"></td>
        <td id="LC627" class="blob-code blob-code-inner js-file-line">				bi.BeginInit()</td>
      </tr>
      <tr>
        <td id="L628" class="blob-num js-line-number" data-line-number="628"></td>
        <td id="LC628" class="blob-code blob-code-inner js-file-line">				bi.StreamSource <span class="pl-k">=</span> fs</td>
      </tr>
      <tr>
        <td id="L629" class="blob-num js-line-number" data-line-number="629"></td>
        <td id="LC629" class="blob-code blob-code-inner js-file-line">				bi.CacheOption <span class="pl-k">=</span> BitmapCacheOption.OnLoad</td>
      </tr>
      <tr>
        <td id="L630" class="blob-num js-line-number" data-line-number="630"></td>
        <td id="LC630" class="blob-code blob-code-inner js-file-line">				bi.CreateOptions <span class="pl-k">=</span> BitmapCreateOptions.None</td>
      </tr>
      <tr>
        <td id="L631" class="blob-num js-line-number" data-line-number="631"></td>
        <td id="LC631" class="blob-code blob-code-inner js-file-line">				bi.EndInit()</td>
      </tr>
      <tr>
        <td id="L632" class="blob-num js-line-number" data-line-number="632"></td>
        <td id="LC632" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L633" class="blob-num js-line-number" data-line-number="633"></td>
        <td id="LC633" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">finally</span>:</td>
      </tr>
      <tr>
        <td id="L634" class="blob-num js-line-number" data-line-number="634"></td>
        <td id="LC634" class="blob-code blob-code-inner js-file-line">				<span class="pl-k">if</span> fs <span class="pl-k">is</span> <span class="pl-k">not</span> <span class="pl-c1">None</span>:</td>
      </tr>
      <tr>
        <td id="L635" class="blob-num js-line-number" data-line-number="635"></td>
        <td id="LC635" class="blob-code blob-code-inner js-file-line">					fs.Close()</td>
      </tr>
      <tr>
        <td id="L636" class="blob-num js-line-number" data-line-number="636"></td>
        <td id="LC636" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L637" class="blob-num js-line-number" data-line-number="637"></td>
        <td id="LC637" class="blob-code blob-code-inner js-file-line">			window.Owner <span class="pl-k">=</span> Application.Current.MainWindow</td>
      </tr>
      <tr>
        <td id="L638" class="blob-num js-line-number" data-line-number="638"></td>
        <td id="LC638" class="blob-code blob-code-inner js-file-line">			window.Title <span class="pl-k">=</span> Application.Current.MainWindow.Title</td>
      </tr>
      <tr>
        <td id="L639" class="blob-num js-line-number" data-line-number="639"></td>
        <td id="LC639" class="blob-code blob-code-inner js-file-line">			window.WindowStartupLocation <span class="pl-k">=</span> WindowStartupLocation.CenterOwner</td>
      </tr>
      <tr>
        <td id="L640" class="blob-num js-line-number" data-line-number="640"></td>
        <td id="LC640" class="blob-code blob-code-inner js-file-line">			window.AllowsTransparency <span class="pl-k">=</span> <span class="pl-c1">True</span></td>
      </tr>
      <tr>
        <td id="L641" class="blob-num js-line-number" data-line-number="641"></td>
        <td id="LC641" class="blob-code blob-code-inner js-file-line">			window.WindowStyle <span class="pl-k">=</span> WindowStyle.None</td>
      </tr>
      <tr>
        <td id="L642" class="blob-num js-line-number" data-line-number="642"></td>
        <td id="LC642" class="blob-code blob-code-inner js-file-line">			window.ResizeMode <span class="pl-k">=</span> ResizeMode.NoResize</td>
      </tr>
      <tr>
        <td id="L643" class="blob-num js-line-number" data-line-number="643"></td>
        <td id="LC643" class="blob-code blob-code-inner js-file-line">			window.ShowActivated <span class="pl-k">=</span> <span class="pl-c1">False</span></td>
      </tr>
      <tr>
        <td id="L644" class="blob-num js-line-number" data-line-number="644"></td>
        <td id="LC644" class="blob-code blob-code-inner js-file-line">			window.ShowInTaskbar <span class="pl-k">=</span> Application.Current.MainWindow.ContextMenu.Items[<span class="pl-c1">5</span>].IsChecked</td>
      </tr>
      <tr>
        <td id="L645" class="blob-num js-line-number" data-line-number="645"></td>
        <td id="LC645" class="blob-code blob-code-inner js-file-line">			window.Topmost <span class="pl-k">=</span> <span class="pl-c1">True</span></td>
      </tr>
      <tr>
        <td id="L646" class="blob-num js-line-number" data-line-number="646"></td>
        <td id="LC646" class="blob-code blob-code-inner js-file-line">			window.SizeToContent <span class="pl-k">=</span> SizeToContent.WidthAndHeight</td>
      </tr>
      <tr>
        <td id="L647" class="blob-num js-line-number" data-line-number="647"></td>
        <td id="LC647" class="blob-code blob-code-inner js-file-line">			window.Background <span class="pl-k">=</span> Brushes.Transparent</td>
      </tr>
      <tr>
        <td id="L648" class="blob-num js-line-number" data-line-number="648"></td>
        <td id="LC648" class="blob-code blob-code-inner js-file-line">			window.Loaded <span class="pl-k">+=</span> onLoaded</td>
      </tr>
      <tr>
        <td id="L649" class="blob-num js-line-number" data-line-number="649"></td>
        <td id="LC649" class="blob-code blob-code-inner js-file-line">			</td>
      </tr>
      <tr>
        <td id="L650" class="blob-num js-line-number" data-line-number="650"></td>
        <td id="LC650" class="blob-code blob-code-inner js-file-line">			contentControl.UseLayoutRounding <span class="pl-k">=</span> <span class="pl-c1">True</span></td>
      </tr>
      <tr>
        <td id="L651" class="blob-num js-line-number" data-line-number="651"></td>
        <td id="LC651" class="blob-code blob-code-inner js-file-line">			contentControl.HorizontalAlignment <span class="pl-k">=</span> HorizontalAlignment.Stretch</td>
      </tr>
      <tr>
        <td id="L652" class="blob-num js-line-number" data-line-number="652"></td>
        <td id="LC652" class="blob-code blob-code-inner js-file-line">			contentControl.VerticalAlignment <span class="pl-k">=</span> VerticalAlignment.Stretch</td>
      </tr>
      <tr>
        <td id="L653" class="blob-num js-line-number" data-line-number="653"></td>
        <td id="LC653" class="blob-code blob-code-inner js-file-line">			contentControl.Opacity <span class="pl-k">=</span> <span class="pl-c1">0</span></td>
      </tr>
      <tr>
        <td id="L654" class="blob-num js-line-number" data-line-number="654"></td>
        <td id="LC654" class="blob-code blob-code-inner js-file-line">			contentControl.RenderTransform <span class="pl-k">=</span> ScaleTransform(<span class="pl-c1">1</span>, <span class="pl-c1">1</span>)</td>
      </tr>
      <tr>
        <td id="L655" class="blob-num js-line-number" data-line-number="655"></td>
        <td id="LC655" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L656" class="blob-num js-line-number" data-line-number="656"></td>
        <td id="LC656" class="blob-code blob-code-inner js-file-line">			window.Content <span class="pl-k">=</span> contentControl</td>
      </tr>
      <tr>
        <td id="L657" class="blob-num js-line-number" data-line-number="657"></td>
        <td id="LC657" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L658" class="blob-num js-line-number" data-line-number="658"></td>
        <td id="LC658" class="blob-code blob-code-inner js-file-line">			imageBrush <span class="pl-k">=</span> ImageBrush(bi)</td>
      </tr>
      <tr>
        <td id="L659" class="blob-num js-line-number" data-line-number="659"></td>
        <td id="LC659" class="blob-code blob-code-inner js-file-line">			imageBrush.TileMode <span class="pl-k">=</span> TileMode.None</td>
      </tr>
      <tr>
        <td id="L660" class="blob-num js-line-number" data-line-number="660"></td>
        <td id="LC660" class="blob-code blob-code-inner js-file-line">			imageBrush.Stretch <span class="pl-k">=</span> Stretch.Fill</td>
      </tr>
      <tr>
        <td id="L661" class="blob-num js-line-number" data-line-number="661"></td>
        <td id="LC661" class="blob-code blob-code-inner js-file-line">			imageBrush.ViewboxUnits <span class="pl-k">=</span> BrushMappingMode.Absolute</td>
      </tr>
      <tr>
        <td id="L662" class="blob-num js-line-number" data-line-number="662"></td>
        <td id="LC662" class="blob-code blob-code-inner js-file-line">			imageBrush.Viewbox <span class="pl-k">=</span> Rect(<span class="pl-c1">0</span>, <span class="pl-c1">0</span>, bi.Width, bi.Height)</td>
      </tr>
      <tr>
        <td id="L663" class="blob-num js-line-number" data-line-number="663"></td>
        <td id="LC663" class="blob-code blob-code-inner js-file-line">			imageBrush.AlignmentX <span class="pl-k">=</span> AlignmentX.Left</td>
      </tr>
      <tr>
        <td id="L664" class="blob-num js-line-number" data-line-number="664"></td>
        <td id="LC664" class="blob-code blob-code-inner js-file-line">			imageBrush.AlignmentY <span class="pl-k">=</span> AlignmentY.Top</td>
      </tr>
      <tr>
        <td id="L665" class="blob-num js-line-number" data-line-number="665"></td>
        <td id="LC665" class="blob-code blob-code-inner js-file-line">			imageBrush.Opacity <span class="pl-k">=</span> <span class="pl-c1">0.5</span></td>
      </tr>
      <tr>
        <td id="L666" class="blob-num js-line-number" data-line-number="666"></td>
        <td id="LC666" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L667" class="blob-num js-line-number" data-line-number="667"></td>
        <td id="LC667" class="blob-code blob-code-inner js-file-line">			dg <span class="pl-k">=</span> DrawingGroup()</td>
      </tr>
      <tr>
        <td id="L668" class="blob-num js-line-number" data-line-number="668"></td>
        <td id="LC668" class="blob-code blob-code-inner js-file-line">			dc <span class="pl-k">=</span> dg.Open()</td>
      </tr>
      <tr>
        <td id="L669" class="blob-num js-line-number" data-line-number="669"></td>
        <td id="LC669" class="blob-code blob-code-inner js-file-line">			dc.DrawRectangle(SolidColorBrush(Color.FromArgb(Byte.MaxValue <span class="pl-k">*</span> <span class="pl-c1">50</span> <span class="pl-k">/</span> <span class="pl-c1">100</span>, <span class="pl-c1">0</span>, <span class="pl-c1">0</span>, <span class="pl-c1">0</span>)), <span class="pl-c1">None</span>, Rect(<span class="pl-c1">0</span>, <span class="pl-c1">0</span>, bi.Width, bi.Height))</td>
      </tr>
      <tr>
        <td id="L670" class="blob-num js-line-number" data-line-number="670"></td>
        <td id="LC670" class="blob-code blob-code-inner js-file-line">			dc.DrawRectangle(imageBrush, <span class="pl-c1">None</span>, Rect(<span class="pl-c1">0</span>, <span class="pl-c1">0</span>, bi.Width, bi.Height))</td>
      </tr>
      <tr>
        <td id="L671" class="blob-num js-line-number" data-line-number="671"></td>
        <td id="LC671" class="blob-code blob-code-inner js-file-line">			dc.Close()</td>
      </tr>
      <tr>
        <td id="L672" class="blob-num js-line-number" data-line-number="672"></td>
        <td id="LC672" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L673" class="blob-num js-line-number" data-line-number="673"></td>
        <td id="LC673" class="blob-code blob-code-inner js-file-line">			backgroundBrush <span class="pl-k">=</span> ImageBrush(DrawingImage(dg))</td>
      </tr>
      <tr>
        <td id="L674" class="blob-num js-line-number" data-line-number="674"></td>
        <td id="LC674" class="blob-code blob-code-inner js-file-line">			backgroundBrush.TileMode <span class="pl-k">=</span> TileMode.Tile</td>
      </tr>
      <tr>
        <td id="L675" class="blob-num js-line-number" data-line-number="675"></td>
        <td id="LC675" class="blob-code blob-code-inner js-file-line">			backgroundBrush.ViewportUnits <span class="pl-k">=</span> BrushMappingMode.Absolute</td>
      </tr>
      <tr>
        <td id="L676" class="blob-num js-line-number" data-line-number="676"></td>
        <td id="LC676" class="blob-code blob-code-inner js-file-line">			backgroundBrush.Viewport <span class="pl-k">=</span> Rect(<span class="pl-c1">0</span>, <span class="pl-c1">0</span>, bi.Width, bi.Height)</td>
      </tr>
      <tr>
        <td id="L677" class="blob-num js-line-number" data-line-number="677"></td>
        <td id="LC677" class="blob-code blob-code-inner js-file-line">			backgroundBrush.Stretch <span class="pl-k">=</span> Stretch.None</td>
      </tr>
      <tr>
        <td id="L678" class="blob-num js-line-number" data-line-number="678"></td>
        <td id="LC678" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L679" class="blob-num js-line-number" data-line-number="679"></td>
        <td id="LC679" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">if</span> backgroundBrush.CanFreeze:</td>
      </tr>
      <tr>
        <td id="L680" class="blob-num js-line-number" data-line-number="680"></td>
        <td id="LC680" class="blob-code blob-code-inner js-file-line">				backgroundBrush.Freeze()</td>
      </tr>
      <tr>
        <td id="L681" class="blob-num js-line-number" data-line-number="681"></td>
        <td id="LC681" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L682" class="blob-num js-line-number" data-line-number="682"></td>
        <td id="LC682" class="blob-code blob-code-inner js-file-line">			grid.HorizontalAlignment <span class="pl-k">=</span> HorizontalAlignment.Stretch</td>
      </tr>
      <tr>
        <td id="L683" class="blob-num js-line-number" data-line-number="683"></td>
        <td id="LC683" class="blob-code blob-code-inner js-file-line">			grid.VerticalAlignment <span class="pl-k">=</span> VerticalAlignment.Stretch</td>
      </tr>
      <tr>
        <td id="L684" class="blob-num js-line-number" data-line-number="684"></td>
        <td id="LC684" class="blob-code blob-code-inner js-file-line">			grid.Background <span class="pl-k">=</span> backgroundBrush</td>
      </tr>
      <tr>
        <td id="L685" class="blob-num js-line-number" data-line-number="685"></td>
        <td id="LC685" class="blob-code blob-code-inner js-file-line">			grid.Width <span class="pl-k">=</span> <span class="pl-c1">150</span></td>
      </tr>
      <tr>
        <td id="L686" class="blob-num js-line-number" data-line-number="686"></td>
        <td id="LC686" class="blob-code blob-code-inner js-file-line">			grid.Height <span class="pl-k">=</span> <span class="pl-c1">150</span></td>
      </tr>
      <tr>
        <td id="L687" class="blob-num js-line-number" data-line-number="687"></td>
        <td id="LC687" class="blob-code blob-code-inner js-file-line">			grid.Clip <span class="pl-k">=</span> EllipseGeometry(Rect(<span class="pl-c1">0</span>, <span class="pl-c1">0</span>, <span class="pl-c1">150</span>, <span class="pl-c1">150</span>))</td>
      </tr>
      <tr>
        <td id="L688" class="blob-num js-line-number" data-line-number="688"></td>
        <td id="LC688" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L689" class="blob-num js-line-number" data-line-number="689"></td>
        <td id="LC689" class="blob-code blob-code-inner js-file-line">			dropShadowEffect <span class="pl-k">=</span> DropShadowEffect()</td>
      </tr>
      <tr>
        <td id="L690" class="blob-num js-line-number" data-line-number="690"></td>
        <td id="LC690" class="blob-code blob-code-inner js-file-line">			dropShadowEffect.BlurRadius <span class="pl-k">=</span> <span class="pl-c1">1</span></td>
      </tr>
      <tr>
        <td id="L691" class="blob-num js-line-number" data-line-number="691"></td>
        <td id="LC691" class="blob-code blob-code-inner js-file-line">			dropShadowEffect.Color <span class="pl-k">=</span> Colors.Black</td>
      </tr>
      <tr>
        <td id="L692" class="blob-num js-line-number" data-line-number="692"></td>
        <td id="LC692" class="blob-code blob-code-inner js-file-line">			dropShadowEffect.Direction <span class="pl-k">=</span> <span class="pl-c1">270</span></td>
      </tr>
      <tr>
        <td id="L693" class="blob-num js-line-number" data-line-number="693"></td>
        <td id="LC693" class="blob-code blob-code-inner js-file-line">			dropShadowEffect.Opacity <span class="pl-k">=</span> <span class="pl-c1">0.5</span></td>
      </tr>
      <tr>
        <td id="L694" class="blob-num js-line-number" data-line-number="694"></td>
        <td id="LC694" class="blob-code blob-code-inner js-file-line">			dropShadowEffect.ShadowDepth <span class="pl-k">=</span> <span class="pl-c1">1</span></td>
      </tr>
      <tr>
        <td id="L695" class="blob-num js-line-number" data-line-number="695"></td>
        <td id="LC695" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L696" class="blob-num js-line-number" data-line-number="696"></td>
        <td id="LC696" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">if</span> dropShadowEffect.CanFreeze:</td>
      </tr>
      <tr>
        <td id="L697" class="blob-num js-line-number" data-line-number="697"></td>
        <td id="LC697" class="blob-code blob-code-inner js-file-line">				dropShadowEffect.Freeze()</td>
      </tr>
      <tr>
        <td id="L698" class="blob-num js-line-number" data-line-number="698"></td>
        <td id="LC698" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L699" class="blob-num js-line-number" data-line-number="699"></td>
        <td id="LC699" class="blob-code blob-code-inner js-file-line">			grid.Effect <span class="pl-k">=</span> dropShadowEffect</td>
      </tr>
      <tr>
        <td id="L700" class="blob-num js-line-number" data-line-number="700"></td>
        <td id="LC700" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L701" class="blob-num js-line-number" data-line-number="701"></td>
        <td id="LC701" class="blob-code blob-code-inner js-file-line">			contentControl.Content <span class="pl-k">=</span> grid</td>
      </tr>
      <tr>
        <td id="L702" class="blob-num js-line-number" data-line-number="702"></td>
        <td id="LC702" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L703" class="blob-num js-line-number" data-line-number="703"></td>
        <td id="LC703" class="blob-code blob-code-inner js-file-line">			solidColorBrush <span class="pl-k">=</span> SolidColorBrush(colorFromAhsb(Byte.MaxValue, <span class="pl-c1">60</span>, <span class="pl-c1">1.0</span>, <span class="pl-c1">1.0</span>) <span class="pl-k">if</span> task.Result.Value.Key <span class="pl-k">&lt;</span> <span class="pl-c1">0</span> <span class="pl-k">else</span> colorFromAhsb(Byte.MaxValue, <span class="pl-c1">0</span>, <span class="pl-c1">1.0</span>, <span class="pl-c1">0.4</span>) <span class="pl-k">if</span> task.Result.Value.Key <span class="pl-k">&gt;</span> <span class="pl-c1">37</span> <span class="pl-k">else</span> colorFromAhsb(Byte.MaxValue, <span class="pl-c1">60</span> <span class="pl-k">-</span> <span class="pl-c1">60</span> <span class="pl-k">*</span> task.Result.Value.Key <span class="pl-k">/</span> <span class="pl-c1">37</span>, <span class="pl-c1">1.0</span>, <span class="pl-c1">0.4</span> <span class="pl-k">+</span> <span class="pl-c1">0.6</span> <span class="pl-k">*</span> Math.Pow(Math.E, (<span class="pl-c1">37</span> <span class="pl-k">/</span> <span class="pl-c1">5</span> <span class="pl-k">-</span> task.Result.Value.Key) <span class="pl-k">-</span> <span class="pl-c1">37</span> <span class="pl-k">/</span> <span class="pl-c1">5</span>) <span class="pl-k">if</span> task.Result.Value.Key <span class="pl-k">&lt;</span> <span class="pl-c1">37</span> <span class="pl-k">/</span> <span class="pl-c1">5</span> <span class="pl-k">else</span> <span class="pl-c1">0.4</span>))</td>
      </tr>
      <tr>
        <td id="L704" class="blob-num js-line-number" data-line-number="704"></td>
        <td id="LC704" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L705" class="blob-num js-line-number" data-line-number="705"></td>
        <td id="LC705" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">if</span> solidColorBrush.CanFreeze:</td>
      </tr>
      <tr>
        <td id="L706" class="blob-num js-line-number" data-line-number="706"></td>
        <td id="LC706" class="blob-code blob-code-inner js-file-line">				solidColorBrush.Freeze()</td>
      </tr>
      <tr>
        <td id="L707" class="blob-num js-line-number" data-line-number="707"></td>
        <td id="LC707" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L708" class="blob-num js-line-number" data-line-number="708"></td>
        <td id="LC708" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L709" class="blob-num js-line-number" data-line-number="709"></td>
        <td id="LC709" class="blob-code blob-code-inner js-file-line">			<span class="pl-k">for</span> stream <span class="pl-k">in</span> task.Result.Value.Value.Value.Value:</td>
      </tr>
      <tr>
        <td id="L710" class="blob-num js-line-number" data-line-number="710"></td>
        <td id="LC710" class="blob-code blob-code-inner js-file-line">				<span class="pl-k">try</span>:</td>
      </tr>
      <tr>
        <td id="L711" class="blob-num js-line-number" data-line-number="711"></td>
        <td id="LC711" class="blob-code blob-code-inner js-file-line">					bi <span class="pl-k">=</span> BitmapImage()</td>
      </tr>
      <tr>
        <td id="L712" class="blob-num js-line-number" data-line-number="712"></td>
        <td id="LC712" class="blob-code blob-code-inner js-file-line">					bi.BeginInit()</td>
      </tr>
      <tr>
        <td id="L713" class="blob-num js-line-number" data-line-number="713"></td>
        <td id="LC713" class="blob-code blob-code-inner js-file-line">					bi.StreamSource <span class="pl-k">=</span> stream</td>
      </tr>
      <tr>
        <td id="L714" class="blob-num js-line-number" data-line-number="714"></td>
        <td id="LC714" class="blob-code blob-code-inner js-file-line">					bi.CacheOption <span class="pl-k">=</span> BitmapCacheOption.OnLoad</td>
      </tr>
      <tr>
        <td id="L715" class="blob-num js-line-number" data-line-number="715"></td>
        <td id="LC715" class="blob-code blob-code-inner js-file-line">					bi.CreateOptions <span class="pl-k">=</span> BitmapCreateOptions.None</td>
      </tr>
      <tr>
        <td id="L716" class="blob-num js-line-number" data-line-number="716"></td>
        <td id="LC716" class="blob-code blob-code-inner js-file-line">					bi.EndInit()</td>
      </tr>
      <tr>
        <td id="L717" class="blob-num js-line-number" data-line-number="717"></td>
        <td id="LC717" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L718" class="blob-num js-line-number" data-line-number="718"></td>
        <td id="LC718" class="blob-code blob-code-inner js-file-line">				<span class="pl-k">finally</span>:</td>
      </tr>
      <tr>
        <td id="L719" class="blob-num js-line-number" data-line-number="719"></td>
        <td id="LC719" class="blob-code blob-code-inner js-file-line">					stream.Close()</td>
      </tr>
      <tr>
        <td id="L720" class="blob-num js-line-number" data-line-number="720"></td>
        <td id="LC720" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L721" class="blob-num js-line-number" data-line-number="721"></td>
        <td id="LC721" class="blob-code blob-code-inner js-file-line">				imageBrush <span class="pl-k">=</span> ImageBrush(bi)</td>
      </tr>
      <tr>
        <td id="L722" class="blob-num js-line-number" data-line-number="722"></td>
        <td id="LC722" class="blob-code blob-code-inner js-file-line">				imageBrush.TileMode <span class="pl-k">=</span> TileMode.None</td>
      </tr>
      <tr>
        <td id="L723" class="blob-num js-line-number" data-line-number="723"></td>
        <td id="LC723" class="blob-code blob-code-inner js-file-line">				imageBrush.ViewportUnits <span class="pl-k">=</span> BrushMappingMode.Absolute</td>
      </tr>
      <tr>
        <td id="L724" class="blob-num js-line-number" data-line-number="724"></td>
        <td id="LC724" class="blob-code blob-code-inner js-file-line">				imageBrush.Viewport <span class="pl-k">=</span> Rect(<span class="pl-c1">0</span>, <span class="pl-c1">0</span>, width, height)</td>
      </tr>
      <tr>
        <td id="L725" class="blob-num js-line-number" data-line-number="725"></td>
        <td id="LC725" class="blob-code blob-code-inner js-file-line">				imageBrush.Stretch <span class="pl-k">=</span> Stretch.Uniform</td>
      </tr>
      <tr>
        <td id="L726" class="blob-num js-line-number" data-line-number="726"></td>
        <td id="LC726" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L727" class="blob-num js-line-number" data-line-number="727"></td>
        <td id="LC727" class="blob-code blob-code-inner js-file-line">				<span class="pl-k">if</span> imageBrush.CanFreeze:</td>
      </tr>
      <tr>
        <td id="L728" class="blob-num js-line-number" data-line-number="728"></td>
        <td id="LC728" class="blob-code blob-code-inner js-file-line">					imageBrush.Freeze()</td>
      </tr>
      <tr>
        <td id="L729" class="blob-num js-line-number" data-line-number="729"></td>
        <td id="LC729" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L730" class="blob-num js-line-number" data-line-number="730"></td>
        <td id="LC730" class="blob-code blob-code-inner js-file-line">				g <span class="pl-k">=</span> Grid()</td>
      </tr>
      <tr>
        <td id="L731" class="blob-num js-line-number" data-line-number="731"></td>
        <td id="LC731" class="blob-code blob-code-inner js-file-line">				g.HorizontalAlignment <span class="pl-k">=</span> HorizontalAlignment.Center</td>
      </tr>
      <tr>
        <td id="L732" class="blob-num js-line-number" data-line-number="732"></td>
        <td id="LC732" class="blob-code blob-code-inner js-file-line">				g.VerticalAlignment <span class="pl-k">=</span> VerticalAlignment.Center</td>
      </tr>
      <tr>
        <td id="L733" class="blob-num js-line-number" data-line-number="733"></td>
        <td id="LC733" class="blob-code blob-code-inner js-file-line">				g.Background <span class="pl-k">=</span> Brushes.Transparent</td>
      </tr>
      <tr>
        <td id="L734" class="blob-num js-line-number" data-line-number="734"></td>
        <td id="LC734" class="blob-code blob-code-inner js-file-line">				g.Width <span class="pl-k">=</span> width</td>
      </tr>
      <tr>
        <td id="L735" class="blob-num js-line-number" data-line-number="735"></td>
        <td id="LC735" class="blob-code blob-code-inner js-file-line">				g.Height <span class="pl-k">=</span> height</td>
      </tr>
      <tr>
        <td id="L736" class="blob-num js-line-number" data-line-number="736"></td>
        <td id="LC736" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L737" class="blob-num js-line-number" data-line-number="737"></td>
        <td id="LC737" class="blob-code blob-code-inner js-file-line">				grid.Children.Add(g)</td>
      </tr>
      <tr>
        <td id="L738" class="blob-num js-line-number" data-line-number="738"></td>
        <td id="LC738" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L739" class="blob-num js-line-number" data-line-number="739"></td>
        <td id="LC739" class="blob-code blob-code-inner js-file-line">				<span class="pl-k">for</span> i <span class="pl-k">in</span> <span class="pl-c1">range</span>(<span class="pl-c1">max</span>):</td>
      </tr>
      <tr>
        <td id="L740" class="blob-num js-line-number" data-line-number="740"></td>
        <td id="LC740" class="blob-code blob-code-inner js-file-line">					rectangle <span class="pl-k">=</span> Rectangle()</td>
      </tr>
      <tr>
        <td id="L741" class="blob-num js-line-number" data-line-number="741"></td>
        <td id="LC741" class="blob-code blob-code-inner js-file-line">					rectangle.HorizontalAlignment <span class="pl-k">=</span> HorizontalAlignment.Left</td>
      </tr>
      <tr>
        <td id="L742" class="blob-num js-line-number" data-line-number="742"></td>
        <td id="LC742" class="blob-code blob-code-inner js-file-line">					rectangle.VerticalAlignment <span class="pl-k">=</span> VerticalAlignment.Top</td>
      </tr>
      <tr>
        <td id="L743" class="blob-num js-line-number" data-line-number="743"></td>
        <td id="LC743" class="blob-code blob-code-inner js-file-line">					rectangle.Width <span class="pl-k">=</span> width</td>
      </tr>
      <tr>
        <td id="L744" class="blob-num js-line-number" data-line-number="744"></td>
        <td id="LC744" class="blob-code blob-code-inner js-file-line">					rectangle.Height <span class="pl-k">=</span> height</td>
      </tr>
      <tr>
        <td id="L745" class="blob-num js-line-number" data-line-number="745"></td>
        <td id="LC745" class="blob-code blob-code-inner js-file-line">					rectangle.Fill <span class="pl-k">=</span> solidColorBrush</td>
      </tr>
      <tr>
        <td id="L746" class="blob-num js-line-number" data-line-number="746"></td>
        <td id="LC746" class="blob-code blob-code-inner js-file-line">					rectangle.Opacity <span class="pl-k">=</span> <span class="pl-c1">0</span></td>
      </tr>
      <tr>
        <td id="L747" class="blob-num js-line-number" data-line-number="747"></td>
        <td id="LC747" class="blob-code blob-code-inner js-file-line">					rectangle.OpacityMask <span class="pl-k">=</span> imageBrush</td>
      </tr>
      <tr>
        <td id="L748" class="blob-num js-line-number" data-line-number="748"></td>
        <td id="LC748" class="blob-code blob-code-inner js-file-line">					rectangle.Clip <span class="pl-k">=</span> RectangleGeometry(Rect(<span class="pl-c1">0</span>, height <span class="pl-k">/</span> <span class="pl-c1">max</span> <span class="pl-k">*</span> i, width, height <span class="pl-k">/</span> <span class="pl-c1">max</span>))</td>
      </tr>
      <tr>
        <td id="L749" class="blob-num js-line-number" data-line-number="749"></td>
        <td id="LC749" class="blob-code blob-code-inner js-file-line">					rectangle.RenderTransform <span class="pl-k">=</span> TranslateTransform(<span class="pl-c1">0</span>, <span class="pl-c1">0</span>)</td>
      </tr>
      <tr>
        <td id="L750" class="blob-num js-line-number" data-line-number="750"></td>
        <td id="LC750" class="blob-code blob-code-inner js-file-line">					rectangle.Tag <span class="pl-k">=</span> i</td>
      </tr>
      <tr>
        <td id="L751" class="blob-num js-line-number" data-line-number="751"></td>
        <td id="LC751" class="blob-code blob-code-inner js-file-line">					</td>
      </tr>
      <tr>
        <td id="L752" class="blob-num js-line-number" data-line-number="752"></td>
        <td id="LC752" class="blob-code blob-code-inner js-file-line">					g.Children.Add(rectangle)</td>
      </tr>
      <tr>
        <td id="L753" class="blob-num js-line-number" data-line-number="753"></td>
        <td id="LC753" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L754" class="blob-num js-line-number" data-line-number="754"></td>
        <td id="LC754" class="blob-code blob-code-inner js-file-line">			window.Show()</td>
      </tr>
      <tr>
        <td id="L755" class="blob-num js-line-number" data-line-number="755"></td>
        <td id="LC755" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L756" class="blob-num js-line-number" data-line-number="756"></td>
        <td id="LC756" class="blob-code blob-code-inner js-file-line">			idList.Clear()</td>
      </tr>
      <tr>
        <td id="L757" class="blob-num js-line-number" data-line-number="757"></td>
        <td id="LC757" class="blob-code blob-code-inner js-file-line">			idList.AddRange(task.Result.Value.Value.Value.Key)</td>
      </tr>
      <tr>
        <td id="L758" class="blob-num js-line-number" data-line-number="758"></td>
        <td id="LC758" class="blob-code blob-code-inner js-file-line">	</td>
      </tr>
      <tr>
        <td id="L759" class="blob-num js-line-number" data-line-number="759"></td>
        <td id="LC759" class="blob-code blob-code-inner js-file-line">	Task.Factory.StartNew[KeyValuePair[List[String], KeyValuePair[Double, KeyValuePair[KeyValuePair[Double, Double], KeyValuePair[List[Double], List[MemoryStream]]]]]](onUpdate, TaskCreationOptions.LongRunning).ContinueWith(Action[Task[KeyValuePair[List[String], KeyValuePair[Double, KeyValuePair[KeyValuePair[Double, Double], KeyValuePair[List[Double], List[MemoryStream]]]]]]](onCompleted), TaskScheduler.FromCurrentSynchronizationContext())</td>
      </tr>
      <tr>
        <td id="L760" class="blob-num js-line-number" data-line-number="760"></td>
        <td id="LC760" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L761" class="blob-num js-line-number" data-line-number="761"></td>
        <td id="LC761" class="blob-code blob-code-inner js-file-line">	timer.Stop()</td>
      </tr>
      <tr>
        <td id="L762" class="blob-num js-line-number" data-line-number="762"></td>
        <td id="LC762" class="blob-code blob-code-inner js-file-line">	timer.Interval <span class="pl-k">=</span> TimeSpan.FromMinutes(<span class="pl-c1">5</span>)</td>
      </tr>
      <tr>
        <td id="L763" class="blob-num js-line-number" data-line-number="763"></td>
        <td id="LC763" class="blob-code blob-code-inner js-file-line">	timer.Start()</td>
      </tr>
      <tr>
        <td id="L764" class="blob-num js-line-number" data-line-number="764"></td>
        <td id="LC764" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L765" class="blob-num js-line-number" data-line-number="765"></td>
        <td id="LC765" class="blob-code blob-code-inner js-file-line"><span class="pl-k">def</span> <span class="pl-en">urlEncode</span>(<span class="pl-smi">value</span>):</td>
      </tr>
      <tr>
        <td id="L766" class="blob-num js-line-number" data-line-number="766"></td>
        <td id="LC766" class="blob-code blob-code-inner js-file-line">	unreserved <span class="pl-k">=</span> <span class="pl-s"><span class="pl-pds">&quot;</span>abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-_.~<span class="pl-pds">&quot;</span></span></td>
      </tr>
      <tr>
        <td id="L767" class="blob-num js-line-number" data-line-number="767"></td>
        <td id="LC767" class="blob-code blob-code-inner js-file-line">	sb <span class="pl-k">=</span> StringBuilder()</td>
      </tr>
      <tr>
        <td id="L768" class="blob-num js-line-number" data-line-number="768"></td>
        <td id="LC768" class="blob-code blob-code-inner js-file-line">	<span class="pl-c1">bytes</span> <span class="pl-k">=</span> Encoding.UTF8.GetBytes(value)</td>
      </tr>
      <tr>
        <td id="L769" class="blob-num js-line-number" data-line-number="769"></td>
        <td id="LC769" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L770" class="blob-num js-line-number" data-line-number="770"></td>
        <td id="LC770" class="blob-code blob-code-inner js-file-line">	<span class="pl-k">for</span> b <span class="pl-k">in</span> <span class="pl-c1">bytes</span>:</td>
      </tr>
      <tr>
        <td id="L771" class="blob-num js-line-number" data-line-number="771"></td>
        <td id="LC771" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">if</span> b <span class="pl-k">&lt;</span> <span class="pl-c1">0x</span>80 <span class="pl-k">and</span> unreserved.IndexOf(Convert.ToChar(b)) <span class="pl-k">!=</span> <span class="pl-k">-</span><span class="pl-c1">1</span>:</td>
      </tr>
      <tr>
        <td id="L772" class="blob-num js-line-number" data-line-number="772"></td>
        <td id="LC772" class="blob-code blob-code-inner js-file-line">			sb.Append(Convert.ToChar(b))</td>
      </tr>
      <tr>
        <td id="L773" class="blob-num js-line-number" data-line-number="773"></td>
        <td id="LC773" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">else</span>:</td>
      </tr>
      <tr>
        <td id="L774" class="blob-num js-line-number" data-line-number="774"></td>
        <td id="LC774" class="blob-code blob-code-inner js-file-line">			sb.Append(<span class="pl-s"><span class="pl-pds">&#39;</span>%<span class="pl-pds">&#39;</span></span> <span class="pl-k">+</span> String.Format(<span class="pl-s"><span class="pl-pds">&quot;</span><span class="pl-c1">{0:X2}</span><span class="pl-pds">&quot;</span></span>, Convert.ToInt32(b)))</td>
      </tr>
      <tr>
        <td id="L775" class="blob-num js-line-number" data-line-number="775"></td>
        <td id="LC775" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L776" class="blob-num js-line-number" data-line-number="776"></td>
        <td id="LC776" class="blob-code blob-code-inner js-file-line">	<span class="pl-k">return</span> sb.ToString()</td>
      </tr>
      <tr>
        <td id="L777" class="blob-num js-line-number" data-line-number="777"></td>
        <td id="LC777" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L778" class="blob-num js-line-number" data-line-number="778"></td>
        <td id="LC778" class="blob-code blob-code-inner js-file-line"><span class="pl-k">def</span> <span class="pl-en">colorFromAhsb</span>(<span class="pl-smi">a</span>, <span class="pl-smi">h</span>, <span class="pl-smi">s</span>, <span class="pl-smi">b</span>):</td>
      </tr>
      <tr>
        <td id="L779" class="blob-num js-line-number" data-line-number="779"></td>
        <td id="LC779" class="blob-code blob-code-inner js-file-line">	<span class="pl-k">if</span> <span class="pl-c1">0</span> <span class="pl-k">&gt;</span> a <span class="pl-k">or</span> <span class="pl-c1">255</span> <span class="pl-k">&lt;</span> a:</td>
      </tr>
      <tr>
        <td id="L780" class="blob-num js-line-number" data-line-number="780"></td>
        <td id="LC780" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">return</span> Colors.Transparent</td>
      </tr>
      <tr>
        <td id="L781" class="blob-num js-line-number" data-line-number="781"></td>
        <td id="LC781" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L782" class="blob-num js-line-number" data-line-number="782"></td>
        <td id="LC782" class="blob-code blob-code-inner js-file-line">	<span class="pl-k">if</span> <span class="pl-c1">0</span> <span class="pl-k">&gt;</span> h <span class="pl-k">or</span> <span class="pl-c1">360</span> <span class="pl-k">&lt;</span> h:</td>
      </tr>
      <tr>
        <td id="L783" class="blob-num js-line-number" data-line-number="783"></td>
        <td id="LC783" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">return</span> Colors.Transparent</td>
      </tr>
      <tr>
        <td id="L784" class="blob-num js-line-number" data-line-number="784"></td>
        <td id="LC784" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L785" class="blob-num js-line-number" data-line-number="785"></td>
        <td id="LC785" class="blob-code blob-code-inner js-file-line">	<span class="pl-k">if</span> <span class="pl-c1">0</span> <span class="pl-k">&gt;</span> s <span class="pl-k">or</span> <span class="pl-c1">1</span> <span class="pl-k">&lt;</span> s:</td>
      </tr>
      <tr>
        <td id="L786" class="blob-num js-line-number" data-line-number="786"></td>
        <td id="LC786" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">return</span> Colors.Transparent</td>
      </tr>
      <tr>
        <td id="L787" class="blob-num js-line-number" data-line-number="787"></td>
        <td id="LC787" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L788" class="blob-num js-line-number" data-line-number="788"></td>
        <td id="LC788" class="blob-code blob-code-inner js-file-line">	<span class="pl-k">if</span> <span class="pl-c1">0</span> <span class="pl-k">&gt;</span> b <span class="pl-k">or</span> <span class="pl-c1">1</span> <span class="pl-k">&lt;</span> b:</td>
      </tr>
      <tr>
        <td id="L789" class="blob-num js-line-number" data-line-number="789"></td>
        <td id="LC789" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">return</span> Colors.Transparent;</td>
      </tr>
      <tr>
        <td id="L790" class="blob-num js-line-number" data-line-number="790"></td>
        <td id="LC790" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L791" class="blob-num js-line-number" data-line-number="791"></td>
        <td id="LC791" class="blob-code blob-code-inner js-file-line">	<span class="pl-k">if</span> <span class="pl-c1">0</span> <span class="pl-k">==</span> s:</td>
      </tr>
      <tr>
        <td id="L792" class="blob-num js-line-number" data-line-number="792"></td>
        <td id="LC792" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">return</span> Color.FromArgb(Convert.ToByte(a), Convert.ToByte(b <span class="pl-k">*</span> Byte.MaxValue), Convert.ToByte(b <span class="pl-k">*</span> Byte.MaxValue), Convert.ToByte(b <span class="pl-k">*</span> Byte.MaxValue))</td>
      </tr>
      <tr>
        <td id="L793" class="blob-num js-line-number" data-line-number="793"></td>
        <td id="LC793" class="blob-code blob-code-inner js-file-line">		</td>
      </tr>
      <tr>
        <td id="L794" class="blob-num js-line-number" data-line-number="794"></td>
        <td id="LC794" class="blob-code blob-code-inner js-file-line">	fMax <span class="pl-k">=</span> b <span class="pl-k">-</span> (b <span class="pl-k">*</span> s) <span class="pl-k">+</span> s <span class="pl-k">if</span> <span class="pl-c1">0.5</span> <span class="pl-k">&lt;</span> b <span class="pl-k">else</span> b <span class="pl-k">+</span> (b <span class="pl-k">*</span> s)</td>
      </tr>
      <tr>
        <td id="L795" class="blob-num js-line-number" data-line-number="795"></td>
        <td id="LC795" class="blob-code blob-code-inner js-file-line">	fMin <span class="pl-k">=</span> b <span class="pl-k">+</span> (b <span class="pl-k">*</span> s) <span class="pl-k">-</span> s <span class="pl-k">if</span> <span class="pl-c1">0.5</span> <span class="pl-k">&lt;</span> b <span class="pl-k">else</span> b <span class="pl-k">-</span> (b <span class="pl-k">*</span> s)</td>
      </tr>
      <tr>
        <td id="L796" class="blob-num js-line-number" data-line-number="796"></td>
        <td id="LC796" class="blob-code blob-code-inner js-file-line">	iSextant <span class="pl-k">=</span> Convert.ToInt32(Math.Floor(h <span class="pl-k">/</span> <span class="pl-c1">60.0</span>))</td>
      </tr>
      <tr>
        <td id="L797" class="blob-num js-line-number" data-line-number="797"></td>
        <td id="LC797" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L798" class="blob-num js-line-number" data-line-number="798"></td>
        <td id="LC798" class="blob-code blob-code-inner js-file-line">	<span class="pl-k">if</span> <span class="pl-c1">300</span> <span class="pl-k">&lt;=</span> h:</td>
      </tr>
      <tr>
        <td id="L799" class="blob-num js-line-number" data-line-number="799"></td>
        <td id="LC799" class="blob-code blob-code-inner js-file-line">		h <span class="pl-k">-=</span> <span class="pl-c1">360</span></td>
      </tr>
      <tr>
        <td id="L800" class="blob-num js-line-number" data-line-number="800"></td>
        <td id="LC800" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L801" class="blob-num js-line-number" data-line-number="801"></td>
        <td id="LC801" class="blob-code blob-code-inner js-file-line">	h <span class="pl-k">=</span> h <span class="pl-k">/</span> <span class="pl-c1">60.0</span></td>
      </tr>
      <tr>
        <td id="L802" class="blob-num js-line-number" data-line-number="802"></td>
        <td id="LC802" class="blob-code blob-code-inner js-file-line">	h <span class="pl-k">-=</span> <span class="pl-c1">2</span> <span class="pl-k">*</span> Convert.ToSingle(Math.Floor(((iSextant <span class="pl-k">+</span> <span class="pl-c1">1</span>) <span class="pl-k">%</span> <span class="pl-c1">6.0</span>) <span class="pl-k">/</span> <span class="pl-c1">2</span>))</td>
      </tr>
      <tr>
        <td id="L803" class="blob-num js-line-number" data-line-number="803"></td>
        <td id="LC803" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L804" class="blob-num js-line-number" data-line-number="804"></td>
        <td id="LC804" class="blob-code blob-code-inner js-file-line">	fMid <span class="pl-k">=</span> h <span class="pl-k">*</span> (fMax <span class="pl-k">-</span> fMin) <span class="pl-k">+</span> fMin <span class="pl-k">if</span> <span class="pl-c1">0</span> <span class="pl-k">==</span> iSextant <span class="pl-k">%</span> <span class="pl-c1">2</span> <span class="pl-k">else</span> fMin <span class="pl-k">-</span> h <span class="pl-k">*</span> (fMax <span class="pl-k">-</span> fMin)</td>
      </tr>
      <tr>
        <td id="L805" class="blob-num js-line-number" data-line-number="805"></td>
        <td id="LC805" class="blob-code blob-code-inner js-file-line">	iMax <span class="pl-k">=</span> Convert.ToInt32(fMax <span class="pl-k">*</span> Byte.MaxValue)</td>
      </tr>
      <tr>
        <td id="L806" class="blob-num js-line-number" data-line-number="806"></td>
        <td id="LC806" class="blob-code blob-code-inner js-file-line">	iMid <span class="pl-k">=</span> Convert.ToInt32(fMid <span class="pl-k">*</span> Byte.MaxValue)</td>
      </tr>
      <tr>
        <td id="L807" class="blob-num js-line-number" data-line-number="807"></td>
        <td id="LC807" class="blob-code blob-code-inner js-file-line">	iMin <span class="pl-k">=</span> Convert.ToInt32(fMin <span class="pl-k">*</span> Byte.MaxValue)</td>
      </tr>
      <tr>
        <td id="L808" class="blob-num js-line-number" data-line-number="808"></td>
        <td id="LC808" class="blob-code blob-code-inner js-file-line">	</td>
      </tr>
      <tr>
        <td id="L809" class="blob-num js-line-number" data-line-number="809"></td>
        <td id="LC809" class="blob-code blob-code-inner js-file-line">	<span class="pl-k">if</span> iSextant <span class="pl-k">==</span> <span class="pl-c1">1</span>:</td>
      </tr>
      <tr>
        <td id="L810" class="blob-num js-line-number" data-line-number="810"></td>
        <td id="LC810" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">return</span> Color.FromArgb(Convert.ToByte(a), Convert.ToByte(iMid), Convert.ToByte(iMax), Convert.ToByte(iMin))</td>
      </tr>
      <tr>
        <td id="L811" class="blob-num js-line-number" data-line-number="811"></td>
        <td id="LC811" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L812" class="blob-num js-line-number" data-line-number="812"></td>
        <td id="LC812" class="blob-code blob-code-inner js-file-line">	<span class="pl-k">elif</span> iSextant <span class="pl-k">==</span> <span class="pl-c1">2</span>:</td>
      </tr>
      <tr>
        <td id="L813" class="blob-num js-line-number" data-line-number="813"></td>
        <td id="LC813" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">return</span> Color.FromArgb(Convert.ToByte(a), Convert.ToByte(iMin), Convert.ToByte(iMax), Convert.ToByte(iMid))</td>
      </tr>
      <tr>
        <td id="L814" class="blob-num js-line-number" data-line-number="814"></td>
        <td id="LC814" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L815" class="blob-num js-line-number" data-line-number="815"></td>
        <td id="LC815" class="blob-code blob-code-inner js-file-line">	<span class="pl-k">elif</span> iSextant <span class="pl-k">==</span> <span class="pl-c1">3</span>:</td>
      </tr>
      <tr>
        <td id="L816" class="blob-num js-line-number" data-line-number="816"></td>
        <td id="LC816" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">return</span> Color.FromArgb(Convert.ToByte(a), Convert.ToByte(iMin), Convert.ToByte(iMid), Convert.ToByte(iMax))</td>
      </tr>
      <tr>
        <td id="L817" class="blob-num js-line-number" data-line-number="817"></td>
        <td id="LC817" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L818" class="blob-num js-line-number" data-line-number="818"></td>
        <td id="LC818" class="blob-code blob-code-inner js-file-line">	<span class="pl-k">elif</span> iSextant <span class="pl-k">==</span> <span class="pl-c1">4</span>:</td>
      </tr>
      <tr>
        <td id="L819" class="blob-num js-line-number" data-line-number="819"></td>
        <td id="LC819" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">return</span> Color.FromArgb(Convert.ToByte(a), Convert.ToByte(iMid), Convert.ToByte(iMin), Convert.ToByte(iMax))</td>
      </tr>
      <tr>
        <td id="L820" class="blob-num js-line-number" data-line-number="820"></td>
        <td id="LC820" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L821" class="blob-num js-line-number" data-line-number="821"></td>
        <td id="LC821" class="blob-code blob-code-inner js-file-line">	<span class="pl-k">elif</span> iSextant <span class="pl-k">==</span> <span class="pl-c1">5</span>:</td>
      </tr>
      <tr>
        <td id="L822" class="blob-num js-line-number" data-line-number="822"></td>
        <td id="LC822" class="blob-code blob-code-inner js-file-line">		<span class="pl-k">return</span> Color.FromArgb(Convert.ToByte(a), Convert.ToByte(iMax), Convert.ToByte(iMin), Convert.ToByte(iMid))</td>
      </tr>
      <tr>
        <td id="L823" class="blob-num js-line-number" data-line-number="823"></td>
        <td id="LC823" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L824" class="blob-num js-line-number" data-line-number="824"></td>
        <td id="LC824" class="blob-code blob-code-inner js-file-line">	<span class="pl-k">return</span> Color.FromArgb(Convert.ToByte(a), Convert.ToByte(iMax), Convert.ToByte(iMid), Convert.ToByte(iMin))</td>
      </tr>
      <tr>
        <td id="L825" class="blob-num js-line-number" data-line-number="825"></td>
        <td id="LC825" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L826" class="blob-num js-line-number" data-line-number="826"></td>
        <td id="LC826" class="blob-code blob-code-inner js-file-line"><span class="pl-k">def</span> <span class="pl-en">onStart</span>(<span class="pl-smi">s</span>, <span class="pl-smi">e</span>):</td>
      </tr>
      <tr>
        <td id="L827" class="blob-num js-line-number" data-line-number="827"></td>
        <td id="LC827" class="blob-code blob-code-inner js-file-line">	<span class="pl-k">global</span> timer</td>
      </tr>
      <tr>
        <td id="L828" class="blob-num js-line-number" data-line-number="828"></td>
        <td id="LC828" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L829" class="blob-num js-line-number" data-line-number="829"></td>
        <td id="LC829" class="blob-code blob-code-inner js-file-line">	timer.Start()</td>
      </tr>
      <tr>
        <td id="L830" class="blob-num js-line-number" data-line-number="830"></td>
        <td id="LC830" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L831" class="blob-num js-line-number" data-line-number="831"></td>
        <td id="LC831" class="blob-code blob-code-inner js-file-line"><span class="pl-k">def</span> <span class="pl-en">onStop</span>(<span class="pl-smi">s</span>, <span class="pl-smi">e</span>):</td>
      </tr>
      <tr>
        <td id="L832" class="blob-num js-line-number" data-line-number="832"></td>
        <td id="LC832" class="blob-code blob-code-inner js-file-line">	<span class="pl-k">global</span> timer</td>
      </tr>
      <tr>
        <td id="L833" class="blob-num js-line-number" data-line-number="833"></td>
        <td id="LC833" class="blob-code blob-code-inner js-file-line">	</td>
      </tr>
      <tr>
        <td id="L834" class="blob-num js-line-number" data-line-number="834"></td>
        <td id="LC834" class="blob-code blob-code-inner js-file-line">	timer.Stop()</td>
      </tr>
      <tr>
        <td id="L835" class="blob-num js-line-number" data-line-number="835"></td>
        <td id="LC835" class="blob-code blob-code-inner js-file-line">
</td>
      </tr>
      <tr>
        <td id="L836" class="blob-num js-line-number" data-line-number="836"></td>
        <td id="LC836" class="blob-code blob-code-inner js-file-line">idList <span class="pl-k">=</span> List[Double]()</td>
      </tr>
      <tr>
        <td id="L837" class="blob-num js-line-number" data-line-number="837"></td>
        <td id="LC837" class="blob-code blob-code-inner js-file-line">timer <span class="pl-k">=</span> DispatcherTimer(DispatcherPriority.Background)</td>
      </tr>
      <tr>
        <td id="L838" class="blob-num js-line-number" data-line-number="838"></td>
        <td id="LC838" class="blob-code blob-code-inner js-file-line">timer.Tick <span class="pl-k">+=</span> onTick</td>
      </tr>
      <tr>
        <td id="L839" class="blob-num js-line-number" data-line-number="839"></td>
        <td id="LC839" class="blob-code blob-code-inner js-file-line">timer.Interval <span class="pl-k">=</span> TimeSpan.FromMilliseconds(<span class="pl-c1">30000</span>)</td>
      </tr>
      <tr>
        <td id="L840" class="blob-num js-line-number" data-line-number="840"></td>
        <td id="LC840" class="blob-code blob-code-inner js-file-line">Script.Instance.Start <span class="pl-k">+=</span> onStart</td>
      </tr>
      <tr>
        <td id="L841" class="blob-num js-line-number" data-line-number="841"></td>
        <td id="LC841" class="blob-code blob-code-inner js-file-line">Script.Instance.Stop <span class="pl-k">+=</span> onStop</td>
      </tr>
</table>

  </div>

</div>

<a href="#jump-to-line" rel="facebox[.linejump]" data-hotkey="l" style="display:none">Jump to Line</a>
<div id="jump-to-line" style="display:none">
  <!-- </textarea> --><!-- '"` --><form accept-charset="UTF-8" action="" class="js-jump-to-line-form" method="get"><div style="margin:0;padding:0;display:inline"><input name="utf8" type="hidden" value="&#x2713;" /></div>
    <input class="linejump-input js-jump-to-line-field" type="text" placeholder="Jump to line&hellip;" aria-label="Jump to line" autofocus>
    <button type="submit" class="btn">Go</button>
</form></div>

  </div>
  <div class="modal-backdrop"></div>
</div>

    </div>
  </div>

    </div>

      <div class="container">
  <div class="site-footer" role="contentinfo">
    <ul class="site-footer-links right">
        <li><a href="https://status.github.com/" data-ga-click="Footer, go to status, text:status">Status</a></li>
      <li><a href="https://developer.github.com" data-ga-click="Footer, go to api, text:api">API</a></li>
      <li><a href="https://training.github.com" data-ga-click="Footer, go to training, text:training">Training</a></li>
      <li><a href="https://shop.github.com" data-ga-click="Footer, go to shop, text:shop">Shop</a></li>
        <li><a href="https://github.com/blog" data-ga-click="Footer, go to blog, text:blog">Blog</a></li>
        <li><a href="https://github.com/about" data-ga-click="Footer, go to about, text:about">About</a></li>
        <li><a href="https://github.com/pricing" data-ga-click="Footer, go to pricing, text:pricing">Pricing</a></li>

    </ul>

    <a href="https://github.com" aria-label="Homepage">
      <span class="mega-octicon octicon-mark-github" title="GitHub"></span>
</a>
    <ul class="site-footer-links">
      <li>&copy; 2015 <span title="0.21232s from github-fe117-cp1-prd.iad.github.net">GitHub</span>, Inc.</li>
        <li><a href="https://github.com/site/terms" data-ga-click="Footer, go to terms, text:terms">Terms</a></li>
        <li><a href="https://github.com/site/privacy" data-ga-click="Footer, go to privacy, text:privacy">Privacy</a></li>
        <li><a href="https://github.com/security" data-ga-click="Footer, go to security, text:security">Security</a></li>
        <li><a href="https://github.com/contact" data-ga-click="Footer, go to contact, text:contact">Contact</a></li>
        <li><a href="https://help.github.com" data-ga-click="Footer, go to help, text:help">Help</a></li>
    </ul>
  </div>
</div>



    
    
    

    <div id="ajax-error-message" class="flash flash-error">
      <span class="octicon octicon-alert"></span>
      <button type="button" class="flash-close js-flash-close js-ajax-error-dismiss" aria-label="Dismiss error">
        <span class="octicon octicon-x"></span>
      </button>
      Something went wrong with that request. Please try again.
    </div>


      <script crossorigin="anonymous" src="https://assets-cdn.github.com/assets/frameworks-b7c9523d93e6cd0235a0a8b7d1a691f7909d3834cd9c9caa7b1674ba508b644c.js"></script>
      <script async="async" crossorigin="anonymous" src="https://assets-cdn.github.com/assets/github-b9083b9659a8922878c9bd553534817e7e081a2cd19d813f7b6c73dce8a87658.js"></script>
      
      
      
    <div class="js-stale-session-flash stale-session-flash flash flash-warn flash-banner hidden">
      <span class="octicon octicon-alert"></span>
      <span class="signed-in-tab-flash">You signed in with another tab or window. <a href="">Reload</a> to refresh your session.</span>
      <span class="signed-out-tab-flash">You signed out in another tab or window. <a href="">Reload</a> to refresh your session.</span>
    </div>
  </body>
</html>
