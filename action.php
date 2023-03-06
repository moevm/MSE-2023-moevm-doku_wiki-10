<?php
/**
 * DokuWiki Plugin xlsx2dw (Action Component)
 *
 * @license GPL 2 http://www.gnu.org/licenses/gpl-2.0.html
 * @author  moevm <ydginster@gmail.com>
 */
class action_plugin_xlsx2dw extends \dokuwiki\Extension\ActionPlugin
{

    /** @inheritDoc */
    public function register(Doku_Event_Handler $controller)
    {
        $controller->register_hook('TOOLBAR_DEFINE', 'FIXME', $this, 'handle_toolbar_define');
   
    }

    /**
     * FIXME Event handler for
     *
     * @param Doku_Event $event  event object by reference
     * @param mixed      $param  optional parameter passed when event was registered
     * @return void
     */
    public function handle_toolbar_define(Doku_Event $event, $param)
    {
    }

}

