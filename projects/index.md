---
layout: page
title: Projects
eyebrow: Projects
lede: >-
  Case studies in port operations, maritime trade, and transportation analytics —
  currently in development.
description: "Case studies in port operations, maritime trade, and transportation analytics, currently in development."
permalink: /projects/
---

<div class="section">
  <div class="container container--narrow">
    {% assign projects = site.projects | sort: "date" | reverse %}
    {% if projects.size > 0 %}
      <div class="grid grid--2">
        {% for project in projects %}
          <a href="{{ project.url | relative_url }}" class="project-card card">
            <div class="project-card__media">
              {% if project.image %}<img src="{{ project.image | relative_url }}" alt="{{ project.title }} — preview">{% endif %}
            </div>
            <p class="project-card__domain">{{ project.domain }}</p>
            <h3>{{ project.title }}</h3>
            <p class="project-card__takeaway">{{ project.summary }}</p>
            <span class="project-card__link">Read case study <svg><use href="#icon-arrow"></use></svg></span>
          </a>
        {% endfor %}
      </div>
    {% else %}
      <div class="project-placeholder">
        <h3 style="margin-bottom: var(--space-4);">New case studies are in progress.</h3>
        <p style="max-width: 56ch; margin-inline: auto; margin-bottom: var(--space-4);">
          I'm building out project work focused on port throughput, maritime trade flows,
          terminal performance, and transportation network analysis — grounded in coursework
          from my Maritime Business Administration &amp; Logistics program and independent
          research. Each one will follow the same standard: a real business question, a clear
          methodology, and a defensible conclusion, not just an exploratory notebook.
        </p>
        <p style="max-width: 56ch; margin-inline: auto;">
          In the meantime, see <a href="{{ '/experience/' | relative_url }}">Experience</a> for
          the analytical and business intelligence work I've already delivered professionally,
          including marine business intelligence at Phillips&nbsp;66.
        </p>
      </div>
    {% endif %}
  </div>
</div>
