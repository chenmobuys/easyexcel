<?php

namespace EasyExcel\Metadata;

class Hyperlink
{
    /**
     * @var string
     */
    private $url;

    /**
     * @var string
     */
    private $tooltip;

    /**
     * Constructor.
     *
     * @param string $url
     * @param string $tooltip
     */
    public function __construct(string $url = '', string $tooltip = '')
    {
        $this->url = $url;

        $this->tooltip = $tooltip;
    }

    /**
     * Get url.
     *
     * @return string
     */
    public function getUrl(): string
    {
        return $this->url;
    }

    /**
     * Set url.
     *
     * @param string $url
     * @return $this
     */
    public function setUrl(string $url): self
    {
        $this->url = $url;

        return $this;
    }

    /**
     * Get tooltip.
     *
     * @return string
     */
    public function getTooltip(): string
    {
        return $this->tooltip;
    }

    /**
     * Set tooltip.
     *
     * @param string $tootip
     *
     * @return $this
     */
    public function setTooltip(string $tootip): self
    {
        $this->tooltip = $tootip;

        return $this;
    }

    /**
     * Is this hyperlink internal? (to another worksheet).
     *
     * @return bool
     */
    public function isInternal(): bool
    {
        return strpos($this->url, 'sheet://') !== false;
    }
}